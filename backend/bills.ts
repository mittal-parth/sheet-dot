import { DurableObject } from "cloudflare:workers";
import {
  createGmailLink, gmailAccountActive, GmailDisconnected, GmailLabelMissing, gmailSearchVersion, isAmount, isDate, isMonth, scanGmail,
  type BillEvent, type Evidence, type GmailAccount, type GmailEnv,
} from "./gmail";
import { parseSchedule, scheduleLabel, type Schedule } from "./watch";

type BillsEnv = GmailEnv & { BILL_WATCHES: DurableObjectNamespace<BillWatch> };
type Periodic = Extract<Schedule, { type: "periodic" }>;
type InputRow = { id: string; merchant: string; month: string; amount: number; due: string; paid: boolean };
type Bill = InputRow & {
  reference?: string | null; notice?: Evidence; payment?: Evidence; review?: Evidence;
  checkedAt?: string; noticeCheckedAt?: string; paidOn?: string | null; termsHistory?: { amount: number; due: string }[];
};
type Watch = {
  id: string; name: string; merchants: string[]; label: string; zone: string; since: string;
  schedule: Periodic; nextAt: number; lastRun?: number; status: string;
  rows: Bill[]; examined: string[]; pendingPayments: BillEvent[]; searchVersion?: number;
};
type State = { account?: GmailAccount; pending?: { account: GmailAccount; expiresAt: number }; watches: Watch[] };
const instruction = "Track new bills for these merchants and reconcile payments.";
const maxRows = 400;
const day = 86_400_000;

function localDate(timestamp: number, zone: string) {
  const parts = new Intl.DateTimeFormat("en", { timeZone: zone, year: "numeric", month: "2-digit", day: "2-digit" })
    .formatToParts(timestamp);
  const part = (type: string) => parts.find(entry => entry.type === type)?.value;
  return `${part("year")}-${part("month")}-${part("day")}`;
}

export function nextBillRun(schedule: Periodic, now: number) {
  const [minute, hour] = schedule.expression.split(" ");
  const nextMinute = (Math.floor(now / 60_000) + 1) * 60_000;
  if (hour === "*") {
    const interval = minute === "*" ? 1 : minute === "0" ? 60 : Number(minute.slice(2));
    return Math.ceil(nextMinute / (interval * 60_000)) * interval * 60_000;
  }
  const local = new Intl.DateTimeFormat("en", { timeZone: schedule.timezone, hour: "numeric", minute: "numeric", hourCycle: "h23" });
  for (let cursor = nextMinute; cursor < nextMinute + 3 * day; cursor += 60_000) {
    const pieces = local.format(cursor).split(":").map(Number);
    if (pieces[0] === Number(hour) && pieces[1] === Number(minute)) return cursor;
  }
  throw new Error("Could not compute next bill check.");
}

export function reconcileBillEvents(watch: Pick<Watch, "rows" | "since" | "pendingPayments">,
  events: BillEvent[], checkedAt: string, newId = () => crypto.randomUUID()) {
  const rows = watch.rows.map(row => ({ ...row }));
  const notices = events.filter(event => event.kind !== "payment").sort((a, b) => a.source.date.localeCompare(b.source.date));
  const candidates = (event: BillEvent) => rows.filter(row => row.merchant === event.merchant && row.month === event.month);
  const byReference = (event: BillEvent, matches: Bill[]) => event.reference
    ? matches.filter(row => row.reference === event.reference) : [];

  for (const notice of notices) {
    const matches = candidates(notice);
    const referenced = byReference(notice, matches);
    const similar = matches.filter(row => (!notice.reference || !row.reference) &&
      (notice.kind === "revision" || row.amount === notice.amount));
    const possible = referenced.length ? referenced : similar;
    const row = possible.length === 1 ? possible[0] : undefined;
    if (!row) {
      if (possible.length > 1 || (matches.length && (!notice.reference || matches.some(match => !match.reference))) ||
        notice.kind === "revision") {
        for (const match of matches) {
          match.review = notice.source;
          match.checkedAt = checkedAt;
        }
        continue;
      }
      if (notice.month >= watch.since && rows.length < maxRows) {
        rows.push({ id: newId(), merchant: notice.merchant, month: notice.month, amount: notice.amount,
          due: notice.due!, paid: false, reference: notice.reference, notice: notice.source, checkedAt, noticeCheckedAt: checkedAt });
      }
      continue;
    }
    if (row.notice && notice.source.date <= row.notice.date) continue;
    const changed = row.amount !== notice.amount || row.due !== notice.due;
    if (changed && (row.paid || notice.kind !== "revision")) {
      row.review = notice.source;
      row.checkedAt = checkedAt;
      continue;
    }
    if (notice.kind === "revision" || !row.notice) {
      if (changed) row.termsHistory = [...(row.termsHistory ?? []), { amount: row.amount, due: row.due }].slice(-12);
      Object.assign(row, { amount: notice.amount, due: notice.due, reference: notice.reference || row.reference,
        notice: notice.source, checkedAt, noticeCheckedAt: checkedAt, review: undefined });
    }
  }

  const pendingPayments: BillEvent[] = [];
  const payments = [...watch.pendingPayments, ...events.filter(event => event.kind === "payment")]
    .sort((a, b) => a.source.date.localeCompare(b.source.date));
  const claimed = new Set(rows.flatMap(row => row.payment ? [row.payment.id] : []));
  for (const receipt of payments) {
    if (claimed.has(receipt.source.id)) continue;
    const matches = candidates(receipt).filter(row => !receipt.reference || !row.reference || row.reference === receipt.reference);
    const referenced = byReference(receipt, matches);
    const exact = (referenced.length ? referenced : matches).filter(row => row.amount === receipt.amount);
    if (exact.length === 1) {
      const row = exact[0];
      if (!row.payment) {
        row.payment = receipt.source;
        row.paidOn = receipt.paidOn;
        row.paid = true;
        row.checkedAt = checkedAt;
        if (row.review?.id === receipt.source.id) row.review = undefined;
      }
      claimed.add(receipt.source.id);
    } else {
      for (const row of exact.length > 1 ? exact : matches) {
        if (!row.paid && row.review?.id !== receipt.source.id) {
          row.review = receipt.source;
          row.checkedAt = checkedAt;
        }
      }
      if (Date.parse(checkedAt) - Date.parse(receipt.source.date) < 90 * day) pendingPayments.push(receipt);
    }
  }
  return { rows, pendingPayments: pendingPayments.slice(-100) };
}

function hasMerchant(value: unknown): value is string {
  return typeof value === "string" && value.length <= 100 && /[\p{L}\p{N}]/u.test(value);
}

function inputRows(value: unknown): value is InputRow[] {
  return Array.isArray(value) && value.length <= maxRows && new Set(value.map(row => row?.id)).size === value.length &&
    value.every(row => row && typeof row.id === "string" && /^[a-f0-9-]{36}$/.test(row.id) &&
      hasMerchant(row.merchant) && isMonth(row.month) && isAmount(row.amount) && isDate(row.due) && typeof row.paid === "boolean");
}

type BillRequest = {
  action: string; id?: string; name?: string; merchants?: string[]; label?: string;
  timezone?: string; schedule?: string; rows?: InputRow[]; callback?: string; importMonth?: string; workbook?: string;
};

export class BillWatch extends DurableObject<BillsEnv> {
  private queue = Promise.resolve<unknown>(undefined);

  private sequential<T>(operation: () => Promise<T>) {
    const next = this.queue.then(operation);
    this.queue = next.catch(() => {});
    return next;
  }

  private state() {
    return this.ctx.storage.kv.get<State>("state") ?? { watches: [] };
  }

  private async save(state: State) {
    this.ctx.storage.kv.put("state", state);
    const next = state.watches.reduce((first, watch) => Math.min(first, watch.nextAt), Infinity);
    if (Number.isFinite(next)) await this.ctx.storage.setAlarm(next);
    else await this.ctx.storage.deleteAlarm();
  }

  async alarm() {
    return this.sequential(async () => {
      const state = this.state();
      const now = Date.now();
      for (const watch of state.watches.filter(watch => watch.nextAt <= now)) {
        watch.nextAt = nextBillRun(watch.schedule, now);
        if (!state.account) {
          watch.status = "Reconnect Gmail";
        } else {
          try {
            const earliest = Math.min(...watch.rows.filter(row => !row.paid).map(row => Date.parse(row.month + "-01")),
              Date.parse(watch.since + "-01"), now - 45 * day);
            const after = Math.max(earliest - 7 * day, now - 366 * day);
            if (watch.searchVersion !== gmailSearchVersion) watch.examined = [];
            const result = await scanGmail(this.env, state.account, watch.merchants, after, watch.label, new Set(watch.examined));
            Object.assign(watch, reconcileBillEvents(watch, result.events, new Date(now).toISOString()));
            watch.examined = [...watch.examined, ...result.examined].slice(-4_000);
            watch.searchVersion = gmailSearchVersion;
            watch.lastRun = Math.floor(now / 1000);
            watch.status = "Active";
            console.info("Bill check:", result.matches, "matching messages,", result.examined.length,
              "examined,", result.events.length, "bill or payment events");
          } catch (error) {
            watch.status = error instanceof GmailDisconnected ? "Reconnect Gmail" :
              error instanceof GmailLabelMissing ? "Gmail label not found" : "Check failed · will retry";
            if (!(error instanceof GmailDisconnected)) console.error("Bill check:", error instanceof Error ? error.message : "Unknown error");
          }
        }
        // Persist each completed table before working on the next one.
        await this.save(state);
      }
      await this.save(state);
    });
  }

  process(input: BillRequest) {
    return this.sequential(async () => {
      const state = this.state();
      if (input.action === "connect") {
        const link = await createGmailLink(this.env, this.ctx.id.toString(), input.callback!);
        state.pending = { account: link.account, expiresAt: Date.now() + 15 * 60_000 };
        await this.save(state);
        return { url: link.url };
      }
      if (input.action === "status") {
        if (state.pending && state.pending.expiresAt <= Date.now()) {
          delete state.pending;
          await this.save(state);
        }
        if (state.pending && await gmailAccountActive(this.env, state.pending.account)) {
          state.account = state.pending.account;
          delete state.pending;
          for (const watch of state.watches) {
            watch.nextAt = Date.now();
            watch.status = "Checking Gmail";
          }
          await this.save(state);
        }
        return {
          connected: !!state.account && !state.watches.some(watch => watch.status === "Reconnect Gmail"),
          connecting: !!state.pending,
          schedules: state.watches.map(watch => ({
            id: watch.id, name: watch.name, question: instruction, schedule: scheduleLabel(watch.schedule),
            status: watch.status, lastRun: watch.lastRun,
          })),
        };
      }

      const existing = state.watches.find(watch => watch.id === input.id);
      if (input.action === "read") return { rows: existing?.rows };
      if (input.action === "stop") {
        state.watches = state.watches.filter(watch => watch !== existing);
        await this.save(state);
        return {};
      }
      if (input.action === "schedule") {
        if (!existing) throw new BillInputError("This bill watch was stopped.");
        const schedule = parseSchedule(input.schedule, existing.zone);
        if (schedule?.type !== "periodic") throw new BillInputError("Use a recurring schedule such as daily 8am or every 1 min.");
        existing.schedule = schedule;
        existing.nextAt = nextBillRun(schedule, Date.now());
        await this.save(state);
        return {};
      }
      if (input.action === "sync") {
        if (!existing) return {};
        for (const row of input.rows!) {
          const saved = existing.rows.find(bill => bill.id === row.id);
          if (!saved && existing.rows.length < maxRows && existing.merchants.includes(row.merchant)) {
            existing.rows.push(row);
          } else if (saved && saved.merchant === row.merchant && saved.month === row.month) {
            const unchanged = saved.amount === row.amount && saved.due === row.due;
            const awaitingSheet = saved.termsHistory?.some(terms => terms.amount === row.amount && terms.due === row.due);
            if (!unchanged && !awaitingSheet) {
              Object.assign(saved, row, { paid: saved.paid || row.paid, notice: undefined, reference: undefined,
                termsHistory: undefined, review: saved.payment ?? saved.review });
            } else if (row.paid) {
              saved.paid = true;
            }
          }
        }
        await this.save(state);
        return {};
      }
      if (input.action !== "start") throw new BillInputError("Unknown bill operation.");
      const zone = new Intl.DateTimeFormat("en", { timeZone: input.timezone }).resolvedOptions().timeZone;
      const schedule = parseSchedule(input.schedule, zone);
      if (schedule?.type !== "periodic") throw new BillInputError("Use a recurring schedule such as daily 8am or every 1 min.");
      const currentMonth = localDate(Date.now(), zone).slice(0, 7);
      const [year, month] = currentMonth.split("-").map(Number);
      const oldestMonth = new Date(Date.UTC(year, month - 12, 1)).toISOString().slice(0, 7);
      if (input.importMonth && (!isMonth(input.importMonth) || input.importMonth < oldestMonth || input.importMonth > currentMonth)) {
        throw new BillInputError("Choose an import month within the last 11 months.");
      }
      if (!state.account) throw new BillInputError("Connect Gmail before starting.");
      if (!existing && state.watches.length >= 5) throw new BillInputError("This workbook already has five bill watches.");
      if (existing) {
        if (existing.label !== (input.label ?? "") || input.importMonth && input.importMonth < existing.since ||
          input.merchants!.some(merchant => !existing.merchants.includes(merchant))) existing.examined = [];
        if (input.importMonth && input.importMonth < existing.since) existing.since = input.importMonth;
        Object.assign(existing, { name: input.name, merchants: input.merchants, label: input.label ?? "", zone,
          schedule, nextAt: Date.now(), status: "Checking Gmail" });
        for (const row of input.rows!) {
          if (!existing.rows.some(bill => bill.id === row.id) && existing.rows.length < maxRows) existing.rows.push(row);
        }
      } else {
        state.watches.push({
          id: input.id!, name: input.name!, merchants: input.merchants!, label: input.label ?? "", zone, schedule,
          nextAt: Date.now(), status: "Checking Gmail", since: input.importMonth || currentMonth,
          rows: input.rows!, examined: [], pendingPayments: [],
        });
      }
      await this.save(state);
      return {};
    });
  }
}

class BillInputError extends Error {
  name = "BillInputError";
}

export async function billRequest(request: Request, env: BillsEnv) {
  const fail = (error: string, status: number) => Response.json({ error }, { status });
  if (request.method !== "POST") return fail("Use POST.", 405);
  if (request.headers.get("Content-Type")?.split(";")[0].trim() !== "application/json") return fail("Expected JSON.", 415);
  const secret = request.headers.get("Authorization")?.match(/^Bearer ([a-f0-9]{64})$/)?.[1];
  if (!secret) return fail("A workbook connection key is required.", 401);
  const raw = await request.text();
  if (raw.length > 170_000) return fail("The bill table is too large.", 413);
  let input: BillRequest;
  try { input = JSON.parse(raw); } catch { return fail("Invalid JSON.", 400); }
  if (!input || !["connect", "status", "read", "stop", "schedule", "sync", "start"].includes(input.action)) {
    return fail("Invalid bill operation.", 400);
  }
  if (typeof input.workbook !== "string" || !/^[a-f0-9]{64}$/.test(input.workbook)) return fail("Invalid workbook.", 400);
  if (!["connect", "status"].includes(input.action) &&
    (typeof input.id !== "string" || !/^[A-Za-z0-9{}_-]{1,128}$/.test(input.id))) return fail("Select an Excel table.", 400);
  if (["start", "sync"].includes(input.action) && !inputRows(input.rows)) return fail("Invalid bill rows.", 400);
  const callbacks = new URL(request.url);
  if (input.action === "connect") input.callback = callbacks.origin + "/gmail-connected.html";
  if (input.action === "start") {
    if (typeof input.name !== "string" || !input.name.trim() || input.name.length > 128 ||
      typeof input.timezone !== "string" || input.timezone.length > 100 ||
      !Array.isArray(input.merchants) || input.merchants.length < 1 || input.merchants.length > 12 ||
      !input.merchants.every(hasMerchant) || new Set(input.merchants).size !== input.merchants.length ||
      (input.label !== undefined && (typeof input.label !== "string" || input.label.length > 100 || /["\\\r\n]/.test(input.label)))) {
      return fail("Choose up to 12 merchants and a valid optional Gmail label.", 400);
    }
    try { new Intl.DateTimeFormat("en", { timeZone: input.timezone }); }
    catch { return fail("Invalid time zone.", 400); }
  }
  try {
    const result = await env.BILL_WATCHES.getByName(input.workbook + ":" + secret).process(input);
    return Response.json(result, { headers: { "Cache-Control": "no-store" } });
  } catch (error) {
    if (error instanceof Error && error.name === "BillInputError") return fail(error.message, 400);
    console.error("Bill watch:", error instanceof Error ? error.message : "Unknown error");
    return fail(error instanceof Error && (error.message.includes("awaiting server configuration") ||
      error.message.startsWith("The Composio project API key lacks permission"))
      ? error.message : "Sheetdot could not complete the bill operation. Try again.", 502);
  }
}
