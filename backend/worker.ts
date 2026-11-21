import { DurableObject } from "cloudflare:workers";
import { answerQuestion, modelInput, type Answer, type AskInput, type Preferences } from "./ask";
import {
  checkRecurringWatch, checkWatch, parseSchedule, scheduleLabel, scheduleRecurringWatch, scheduleWatch,
  stopWatch, updateWatchPreferences, updateWatchSchedule, type SavedWatch, type Schedule,
} from "./watch";

type Env = {
  WORKBOOKS: DurableObjectNamespace<Workbook>;
  OPENAI_API_KEY?: string;
};

type WatchEntry = {
  input: AskInput;
  original: Schedule;
  schedule: Schedule;
  watch: SavedWatch;
  aboutRevision?: number;
  removed?: boolean;
};

type AboutMe = { preferences: Preferences; revision: number };

function watchAddress(address: unknown): string | undefined {
  if (typeof address !== "string" || address.length > 256) return;
  const separator = address.lastIndexOf("!");
  const cell = address.slice(separator + 1).replaceAll("$", "").toUpperCase();
  if (separator < 1 || !/^[A-Z]{1,3}[1-9]\d*$/.test(cell)) return;
  const sheet = address.slice(0, separator).replace(/^'(.*)'$/, "$1").replaceAll("''", "'");
  return `${sheet.toUpperCase()}!${cell}`;
}

export class Workbook extends DurableObject<Env> {
  private pending = new Map<string, Promise<Answer>>();
  private watchOperations = new Map<string, Promise<unknown>>();

  ping() {
    return this.ctx.storage.sql.exec<{ value: number }>("SELECT 1 AS value").one().value;
  }

  async ask(input: AskInput, preferences = this.aboutMe().preferences): Promise<{ answer: Answer; cached: boolean }> {
    const sql = this.ctx.storage.sql;
    sql.exec("CREATE TABLE IF NOT EXISTS answers (input TEXT PRIMARY KEY, answer TEXT NOT NULL)");
    const key = JSON.stringify(modelInput(input, preferences));
    const saved = sql.exec<{ answer: string }>(
      "SELECT answer FROM answers WHERE input = ?", key
    ).toArray()[0];
    if (saved) return { answer: JSON.parse(saved.answer), cached: true };

    const pending = this.pending.get(key);
    if (pending) return { answer: await pending, cached: true };

    const apiKey = this.env.OPENAI_API_KEY;
    if (!apiKey) throw new Error("OPENAI_API_KEY is missing from the Worker environment");
    const request = answerQuestion(apiKey, input, preferences).then(answer => {
      sql.exec("INSERT INTO answers (input, answer) VALUES (?, ?)", key, JSON.stringify(answer));
      return answer;
    });
    this.pending.set(key, request);
    try {
      return { answer: await request, cached: false };
    } finally {
      this.pending.delete(key);
    }
  }

  private async serializeWatch<T>(address: string, operation: () => Promise<T>): Promise<T> {
    const previous = this.watchOperations.get(address);
    const next = (previous ?? Promise.resolve()).catch(() => undefined).then(operation);
    this.watchOperations.set(address, next);
    try {
      return await next;
    } finally {
      if (this.watchOperations.get(address) === next) this.watchOperations.delete(address);
    }
  }

  private watchStorage() {
    const sql = this.ctx.storage.sql;
    this.ctx.storage.transactionSync(() => {
      const tables = new Set(sql.exec<{ name: string }>(
        "SELECT name FROM sqlite_master WHERE type = 'table'"
      ).toArray().map(row => row.name));
      if (tables.has("watch_entries")) return;
      sql.exec("CREATE TABLE watch_entries (address TEXT PRIMARY KEY, entry TEXT NOT NULL)");
      if (tables.has("watches")) {
        const schedule = { type: "once" } as const;
        for (const row of sql.exec<{ input: string; watch: string }>("SELECT input, watch FROM watches ORDER BY rowid")) {
          const [source, input] = JSON.parse(row.input);
          const address = watchAddress(source);
          if (address) this.saveWatch(address, { input, original: schedule, schedule, watch: JSON.parse(row.watch) });
        }
      }
      if (tables.has("recurring_watches")) {
        for (const row of sql.exec<{ address: string; definition: string; watch: string }>(
          "SELECT address, definition, watch FROM recurring_watches"
        )) {
          const [input, schedule] = JSON.parse(row.definition);
          this.saveWatch(row.address, { input, original: schedule, schedule, watch: JSON.parse(row.watch) });
        }
      }
    });
    return sql;
  }

  private readWatch(address: string) {
    const row = this.watchStorage().exec<{ entry: string }>(
      "SELECT entry FROM watch_entries WHERE address = ?", address
    ).toArray()[0];
    if (row) return JSON.parse(row.entry) as WatchEntry;
  }

  private saveWatch(address: string, entry: WatchEntry) {
    this.ctx.storage.sql.exec(
      "INSERT OR REPLACE INTO watch_entries (address, entry) VALUES (?, ?)",
      address, JSON.stringify(entry)
    );
  }

  private watchKey() {
    const key = this.env.OPENAI_API_KEY;
    if (!key) throw new Error("OPENAI_API_KEY is missing from the Worker environment");
    return key;
  }

  private aboutMe(): AboutMe {
    return this.ctx.storage.kv.get<AboutMe>("about-me") ?? { preferences: [], revision: 0 };
  }

  async saveAboutMe(preferences: Preferences) {
    const saved = this.aboutMe();
    if (JSON.stringify(saved.preferences) !== JSON.stringify(preferences)) {
      this.ctx.storage.kv.put("about-me", { preferences, revision: saved.revision + 1 });
    }
    const entries = this.watchStorage().exec<{ address: string; entry: string }>(
      "SELECT address, entry FROM watch_entries"
    ).toArray();
    const revision = this.aboutMe().revision;
    const updates = entries
      .filter(row => this.watchNeedsPreferences(JSON.parse(row.entry), revision))
      .map(({ address }) => this.serializeWatch(address, async () => {
        const entry = this.readWatch(address);
        if (entry) await this.syncWatchPreferences(address, entry);
      }));
    const results = await Promise.allSettled(updates);
    const failed = results.find(result => result.status === "rejected");
    if (failed) throw failed.reason;
  }

  private watchNeedsPreferences(entry: WatchEntry, revision: number) {
    if (entry.removed || (entry.aboutRevision ?? 0) === revision) return false;
    const { watch } = entry;
    return "recurring" in watch
      ? watch.status !== "stopped"
      : watch.status === "scheduled" && watch.runAt > Date.now() / 1000;
  }

  private async syncWatchPreferences(address: string, entry: WatchEntry) {
    const about = this.aboutMe();
    if (!this.watchNeedsPreferences(entry, about.revision)) return entry;
    await updateWatchPreferences(this.watchKey(), entry.watch.id, entry.input, about.preferences);
    const updated = { ...entry, aboutRevision: about.revision };
    this.saveWatch(address, updated);
    return updated;
  }

  private async retireWatch(address: string, entry: WatchEntry, removed = false) {
    const watch = entry.watch.status === "scheduled" || ("recurring" in entry.watch && entry.watch.status === "complete")
      ? await stopWatch(this.watchKey(), entry.watch)
      : entry.watch;
    this.saveWatch(address, { ...entry, watch, removed });
    return watch;
  }

  private async startWatch(input: AskInput, schedule: Schedule) {
    const key = this.watchKey();
    const { preferences, revision } = this.aboutMe();
    const watch = await (schedule.type === "once"
      ? scheduleWatch(key, input, preferences)
      : scheduleRecurringWatch(key, input, schedule, preferences));
    return { watch, aboutRevision: revision };
  }

  listWatches() {
    return this.watchStorage().exec<{ address: string; entry: string }>(
      "SELECT address, entry FROM watch_entries ORDER BY address"
    ).toArray().flatMap(({ address, entry }) => {
      const saved: WatchEntry = JSON.parse(entry);
      if (saved.removed) return [];
      const { watch } = saved;
      const status = "recurring" in watch && watch.status !== "stopped" ? "Active" : {
        scheduled: "Scheduled", complete: "Complete", stopped: "Stopped", failed: "Failed",
      }[watch.status];
      const lastRun = "recurring" in watch ? watch.lastRunAt : watch.status === "complete" ? watch.runAt : undefined;
      return [{ address, question: saved.input.question, schedule: scheduleLabel(saved.schedule), status, lastRun }];
    });
  }

  stopWatch(address: string, remove = false) {
    return this.serializeWatch(address, async () => {
      const entry = this.readWatch(address);
      if (!entry || entry.removed || (!remove && entry.watch.status === "stopped")) return false;
      await this.retireWatch(address, entry, remove);
      return true;
    });
  }

  editWatch(address: string, label: unknown, timezone: string) {
    return this.serializeWatch(address, async () => {
      const entry = this.readWatch(address);
      if (!entry || entry.removed) return "This WATCH was already removed.";
      const zone = entry.schedule.type === "periodic" && entry.schedule.timezone !== "UTC"
        ? entry.schedule.timezone : timezone;
      const schedule = parseSchedule(label, zone);
      if (!schedule) return "Use a supported schedule such as every 2 min or daily 8am.";
      if (JSON.stringify(schedule) === JSON.stringify(entry.schedule) && entry.watch.status !== "stopped") return;
      let updated;
      if ("recurring" in entry.watch && entry.watch.status !== "stopped" && schedule.type === "periodic" &&
        entry.schedule.type === "periodic" && entry.schedule.timezone === schedule.timezone) {
        const watch = await updateWatchSchedule(this.watchKey(), entry.watch, schedule);
        updated = { ...entry, schedule, watch };
      } else {
        await this.retireWatch(address, entry);
        updated = { ...entry, schedule, ...await this.startWatch(entry.input, schedule) };
      }
      this.saveWatch(address, updated);
      await this.syncWatchPreferences(address, updated);
    });
  }

  watch(address: string, input: AskInput, original: Schedule) {
    return this.serializeWatch(address, async () => {
      let entry = this.readWatch(address);
      const sameSchedule = entry && JSON.stringify(entry.original) === JSON.stringify(original);
      if (sameSchedule && (entry.removed || entry.watch.status === "stopped")) return entry.watch;

      if (sameSchedule && JSON.stringify(entry.input) === JSON.stringify(input)) {
        entry = await this.syncWatchPreferences(address, entry);
        const { watch } = entry;
        if ("recurring" in watch) {
          if (Date.now() < watch.checkAfter) return watch;
        } else if (watch.status !== "scheduled" || Date.now() / 1000 < watch.runAt) {
          return watch;
        }
        const result = "recurring" in watch
          ? await checkRecurringWatch(this.watchKey(), watch)
          : await checkWatch(this.watchKey(), watch);
        if (result !== watch) this.saveWatch(address, { ...entry, watch: result });
        return result;
      }

      const schedule = sameSchedule ? entry.schedule : original;
      if (entry && !entry.removed) await this.retireWatch(address, entry);
      const created = { input, original, schedule, ...await this.startWatch(input, schedule) };
      this.saveWatch(address, created);
      return (await this.syncWatchPreferences(address, created)).watch;
    });
  }
}

export default {
  async fetch(request: Request, env: Env) {
    const { pathname } = new URL(request.url);
    if (pathname === "/health") {
      const workbook = env.WORKBOOKS.getByName("compatibility-test");
      return Response.json({ sqlite: await workbook.ping() === 1 });
    }

    if (!["/ask", "/watch", "/automations", "/about-me"].includes(pathname)) return new Response("Unknown route", { status: 404 });
    const stopping = pathname === "/watch" && request.method === "DELETE";
    const managing = pathname === "/automations" && ["PATCH", "DELETE"].includes(request.method);
    if (request.method !== "POST" && !stopping && !managing) {
      const allow = pathname === "/automations" ? "POST, PATCH, DELETE" : pathname === "/watch" ? "POST, DELETE" : "POST";
      return new Response("Method not allowed", { status: 405, headers: { Allow: allow } });
    }
    if (request.headers.get("Content-Type")?.split(";")[0].trim() !== "application/json") {
      return Response.json({ error: "Expected application/json." }, { status: 415 });
    }

    const raw = await request.text();
    if (raw.length > 32_000) {
      return Response.json({ error: "Question and row are too large." }, { status: 413 });
    }

    let body;
    try {
      body = JSON.parse(raw);
    } catch {
      return Response.json({ error: "Invalid JSON." }, { status: 400 });
    }
    const { workbook, question, row = [], schedule, address, timezone = "UTC", preferences } = body ?? {};
    const cell = watchAddress(address);
    if (typeof workbook !== "string" || !/^[a-f0-9]{64}$/.test(workbook) ||
      ((pathname === "/watch" || managing) && !cell)) {
      return Response.json({ error: "Expected a workbook ID and a worksheet cell for WATCH changes." }, { status: 400 });
    }
    const target = env.WORKBOOKS.getByName(workbook);
    if ((pathname === "/about-me" || (pathname === "/ask" && preferences !== undefined)) &&
      (!Array.isArray(preferences) || preferences.length > 100 ||
        !preferences.every(pair => Array.isArray(pair) && pair.length === 2 &&
          pair.every(value => typeof value === "string" && value.trim())))) {
      return Response.json({ error: "Expected up to 100 About me preferences with a name and value." }, { status: 400 });
    }
    if (pathname === "/about-me") {
      try {
        await target.saveAboutMe(preferences);
        return Response.json({ ok: true });
      } catch (error) {
        console.error(error instanceof Error ? error.message : "Unknown preferences error");
        return Response.json({ error: "Could not sync About me preferences to WATCH. Sheetdot will retry." }, { status: 502 });
      }
    }
    if (pathname === "/automations") {
      try {
        if (!managing) return Response.json({ schedules: await target.listWatches() });
        if (request.method === "DELETE") {
          await target.stopWatch(cell!, true);
        } else {
          const error = await target.editWatch(cell!, schedule, typeof timezone === "string" ? timezone : "UTC");
          if (error) return Response.json({ error }, { status: 400 });
        }
        return Response.json({ ok: true });
      } catch (error) {
        console.error(error instanceof Error ? error.message : "Unknown automation error");
        return Response.json({ error: "Could not save the WATCH schedule change." }, { status: 502 });
      }
    }
    if (stopping && cell) {
      try {
        return Response.json({ stopped: await target.stopWatch(cell) });
      } catch (error) {
        console.error(error instanceof Error ? error.message : "Unknown OpenAI error");
        return Response.json({ error: "Sheetdot could not stop the schedule. Try again." }, { status: 502 });
      }
    }
    if (
      typeof question !== "string" || !question.trim() ||
      !Array.isArray(row) || !row.every(values => Array.isArray(values) && values.every(
        value => value === null || ["string", "number", "boolean"].includes(typeof value)
      ))
    ) {
      return Response.json({ error: "Expected a question and optional range." }, { status: 400 });
    }

    try {
      const input = { question: question.trim(), row };
      if (pathname === "/ask") return Response.json(await target.ask(input, preferences));

      const timing = parseSchedule(schedule, timezone);
      if (!timing) {
        return Response.json({ error: 'Use "in 1 min", a minute interval that divides an hour, or a daily time like "daily 8am".' }, { status: 400 });
      }
      return Response.json(await target.watch(cell!, input, timing));
    } catch (error) {
      console.error(error instanceof Error ? error.message : "Unknown OpenAI error");
      const operation = pathname === "/watch" ? "check the schedule" : "get an OpenAI answer";
      return Response.json({ error: `Sheetdot could not ${operation}. Recalculate to retry.` }, { status: 502 });
    }
  }
};
