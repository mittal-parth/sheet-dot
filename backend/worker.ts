import { DurableObject } from "cloudflare:workers";
import { answerQuestion, type Answer, type AskInput } from "./ask";
import {
  checkRecurringWatch, checkWatch, parseSchedule, scheduleRecurringWatch, scheduleWatch,
  stopRecurringWatch, type RecurringWatch, type Schedule, type Watch,
} from "./watch";

type Env = {
  WORKBOOKS: DurableObjectNamespace<Workbook>;
  OPENAI_API_KEY?: string;
};

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

  async ask(input: AskInput): Promise<{ answer: Answer; cached: boolean }> {
    const sql = this.ctx.storage.sql;
    sql.exec("CREATE TABLE IF NOT EXISTS answers (input TEXT PRIMARY KEY, answer TEXT NOT NULL)");
    const key = JSON.stringify(input);
    const saved = sql.exec<{ answer: string }>(
      "SELECT answer FROM answers WHERE input = ?", key
    ).toArray()[0];
    if (saved) return { answer: JSON.parse(saved.answer), cached: true };

    const pending = this.pending.get(key);
    if (pending) return { answer: await pending, cached: true };

    const apiKey = this.env.OPENAI_API_KEY;
    if (!apiKey) throw new Error("OPENAI_API_KEY is missing from the Worker environment");
    const request = answerQuestion(apiKey, input).then(answer => {
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

  private readRecurring(address: string) {
    const sql = this.ctx.storage.sql;
    sql.exec("CREATE TABLE IF NOT EXISTS recurring_watches (address TEXT PRIMARY KEY, definition TEXT NOT NULL, watch TEXT NOT NULL)");
    const row = sql.exec<{ definition: string; watch: string }>(
      "SELECT definition, watch FROM recurring_watches WHERE address = ?", address
    ).toArray()[0];
    if (row) return { definition: row.definition, watch: JSON.parse(row.watch) as RecurringWatch };
  }

  private saveRecurring(address: string, definition: string, watch: RecurringWatch) {
    this.ctx.storage.sql.exec(
      "INSERT OR REPLACE INTO recurring_watches (address, definition, watch) VALUES (?, ?, ?)",
      address, definition, JSON.stringify(watch)
    );
  }

  private async stopSavedWatch(address: string) {
    const saved = this.readRecurring(address);
    if (!saved || saved.watch.status === "stopped") return false;

    const apiKey = this.env.OPENAI_API_KEY;
    if (!apiKey) throw new Error("OPENAI_API_KEY is missing from the Worker environment");
    const stopped = await stopRecurringWatch(apiKey, saved.watch);
    this.saveRecurring(address, saved.definition, stopped);
    return true;
  }

  stopWatch(address: string) {
    return this.serializeWatch(address, () => this.stopSavedWatch(address));
  }

  watch(address: string, originalAddress: string, input: AskInput, schedule: Schedule) {
    return this.serializeWatch(address, async () => {
      if (schedule.type === "once") {
        await this.stopSavedWatch(address);
        return this.watchOnce(originalAddress, input);
      }

      const saved = this.readRecurring(address);
      const definition = JSON.stringify([input, schedule]);
      if (saved?.definition === definition &&
        (saved.watch.status === "stopped" || Date.now() < saved.watch.checkAfter)) {
        return saved.watch;
      }

      const apiKey = this.env.OPENAI_API_KEY;
      if (!apiKey) throw new Error("OPENAI_API_KEY is missing from the Worker environment");

      let watch;
      if (saved?.definition === definition) {
        watch = await checkRecurringWatch(apiKey, saved.watch);
      } else {
        await this.stopSavedWatch(address);
        watch = await scheduleRecurringWatch(apiKey, input, schedule);
      }
      this.saveRecurring(address, definition, watch);
      return watch;
    });
  }

  private async watchOnce(address: string, input: AskInput): Promise<Watch> {
    const sql = this.ctx.storage.sql;
    sql.exec("CREATE TABLE IF NOT EXISTS watches (input TEXT PRIMARY KEY, watch TEXT NOT NULL)");
    const key = JSON.stringify([address, input]);
    const saved = sql.exec<{ watch: string }>(
      "SELECT watch FROM watches WHERE input = ?", key
    ).toArray()[0];
    const watch: Watch | undefined = saved && JSON.parse(saved.watch);
    if (watch && (watch.status !== "scheduled" || Date.now() / 1000 < watch.runAt)) return watch;

    const apiKey = this.env.OPENAI_API_KEY;
    if (!apiKey) throw new Error("OPENAI_API_KEY is missing from the Worker environment");
    const updated = await (watch ? checkWatch(apiKey, watch) : scheduleWatch(apiKey, input));
    if (updated !== watch) {
      sql.exec("INSERT OR REPLACE INTO watches (input, watch) VALUES (?, ?)", key, JSON.stringify(updated));
    }
    return updated;
  }
}

export default {
  async fetch(request: Request, env: Env) {
    const { pathname } = new URL(request.url);
    if (pathname === "/health") {
      const workbook = env.WORKBOOKS.getByName("compatibility-test");
      return Response.json({ sqlite: await workbook.ping() === 1 });
    }

    if (pathname !== "/ask" && pathname !== "/watch") return new Response("Unknown route", { status: 404 });
    const stopping = pathname === "/watch" && request.method === "DELETE";
    if (request.method !== "POST" && !stopping) {
      const allow = pathname === "/watch" ? "POST, DELETE" : "POST";
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
    const { workbook, question, row = [], schedule, address, timezone = "UTC" } = body ?? {};
    const cell = pathname === "/watch" ? watchAddress(address) : undefined;
    if (typeof workbook !== "string" || !/^[a-f0-9]{64}$/.test(workbook) || (pathname === "/watch" && !cell)) {
      return Response.json({ error: "Expected a workbook ID and a worksheet cell for WATCH." }, { status: 400 });
    }
    const target = env.WORKBOOKS.getByName(workbook);
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
      if (!cell) return Response.json(await target.ask(input));

      const timing = parseSchedule(schedule, timezone);
      if (!timing) {
        return Response.json({ error: 'Use "in 1 min", a minute interval that divides an hour, or a daily time like "daily 8am".' }, { status: 400 });
      }
      return Response.json(await target.watch(cell, address, input, timing));
    } catch (error) {
      console.error(error instanceof Error ? error.message : "Unknown OpenAI error");
      const operation = pathname === "/watch" ? "check the schedule" : "get an OpenAI answer";
      return Response.json({ error: `Sheetdot could not ${operation}. Recalculate to retry.` }, { status: 502 });
    }
  }
};
