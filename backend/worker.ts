import { DurableObject } from "cloudflare:workers";
import { answerQuestion, type Answer, type AskInput } from "./ask";
import { checkWatch, scheduleWatch, type Watch } from "./watch";

type Env = {
  WORKBOOKS: DurableObjectNamespace<Workbook>;
  OPENAI_API_KEY?: string;
};

export class Workbook extends DurableObject<Env> {
  private pending = new Map<string, Promise<Answer>>();
  private pendingWatches = new Map<string, Promise<Watch>>();

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

  async watch(address: string, input: AskInput): Promise<Watch> {
    const sql = this.ctx.storage.sql;
    sql.exec("CREATE TABLE IF NOT EXISTS watches (input TEXT PRIMARY KEY, watch TEXT NOT NULL)");
    const key = JSON.stringify([address, input]);
    const saved = sql.exec<{ watch: string }>(
      "SELECT watch FROM watches WHERE input = ?", key
    ).toArray()[0];
    const watch: Watch | undefined = saved && JSON.parse(saved.watch);
    if (watch && (watch.status !== "scheduled" || Date.now() / 1000 < watch.runAt)) return watch;

    const pending = this.pendingWatches.get(key);
    if (pending) return pending;

    const apiKey = this.env.OPENAI_API_KEY;
    if (!apiKey) throw new Error("OPENAI_API_KEY is missing from the Worker environment");
    const request = (watch ? checkWatch(apiKey, watch) : scheduleWatch(apiKey, input)).then(updated => {
      if (updated !== watch) {
        sql.exec("INSERT OR REPLACE INTO watches (input, watch) VALUES (?, ?)", key, JSON.stringify(updated));
      }
      return updated;
    });
    this.pendingWatches.set(key, request);
    try {
      return await request;
    } finally {
      this.pendingWatches.delete(key);
    }
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
    if (request.method !== "POST") {
      return new Response("Method not allowed", { status: 405, headers: { Allow: "POST" } });
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
    const { workbook, question, row = [], schedule, address } = body ?? {};
    if (
      typeof workbook !== "string" || !/^[a-f0-9]{64}$/.test(workbook) ||
      typeof question !== "string" || !question.trim() ||
      !Array.isArray(row) || !row.every(values => Array.isArray(values) && values.every(
        value => value === null || ["string", "number", "boolean"].includes(typeof value)
      ))
    ) {
      return Response.json({ error: "Expected a workbook ID, question and optional range." }, { status: 400 });
    }
    if (pathname === "/watch" && (schedule !== "in 1 min" || typeof address !== "string" || !address || address.length > 256)) {
      return Response.json({ error: 'WATCH currently supports the one-time schedule "in 1 min".' }, { status: 400 });
    }

    try {
      const target = env.WORKBOOKS.getByName(workbook);
      const input = { question: question.trim(), row };
      return Response.json(await (pathname === "/watch" ? target.watch(address, input) : target.ask(input)));
    } catch (error) {
      console.error(error instanceof Error ? error.message : "Unknown OpenAI error");
      const operation = pathname === "/watch" ? "check the schedule" : "get an OpenAI answer";
      return Response.json({ error: `Sheetdot could not ${operation}. Recalculate to retry.` }, { status: 502 });
    }
  }
};
