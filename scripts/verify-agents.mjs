import { randomUUID } from "node:crypto";
import { parseArgs } from "node:util";
import { setTimeout as delay } from "node:timers/promises";

const { values } = parseArgs({ options: { schedule: { type: "boolean", default: false } } });
const apiKey = process.env.OPENAI_API_KEY;
const automations = "/v1/agents/automations";
const model = process.env.SHEETDOT_AGENT_MODEL ?? "gpt-6-sol";
const stop = new AbortController();

async function request(path, method = "GET", body) {
  let response;
  try {
    response = await fetch(new URL(path, "https://api.openai.com"), {
      method,
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "OpenAI-Beta": "agents=v1",
        ...(body && { "Content-Type": "application/json" }),
      },
      body: body ? JSON.stringify(body) : undefined,
      redirect: "error",
      signal: AbortSignal.timeout(15_000),
    });
  } catch (error) {
    const reason = error.cause?.code || error.cause?.name || error.name;
    throw new Error(`${method} ${path} could not reach api.openai.com (${reason})`);
  }
  if (response.status === 204) return;
  const data = await response.json().catch(() => null);
  if (!response.ok) {
    const code = data?.error?.code ?? data?.error?.type;
    throw new Error(`${method} ${path} returned HTTP ${response.status}${code ? ` (${code})` : ""}`);
  }
  return data;
}

async function waitFor(label, check, timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    stop.signal.throwIfAborted();
    const result = await check();
    if (result) return result;
    await delay(5_000, undefined, { signal: stop.signal });
  }
  throw new Error(`Timed out waiting for ${label}`);
}

async function verifySchedule() {
  const name = `Sheetdot schedule verification ${randomUUID()}`;
  const token = "Sheetdot scheduled run verified";
  console.log(`Creating ${name} with cron * * * * * in UTC`);
  const schedule = await request(automations, "POST", {
    name,
    type: "periodic",
    expression: "* * * * *",
    timezone: "UTC",
    target: {
      type: "new_session",
      agent: {
        model,
        instructions: "Use web search for the requested lookup. Return a single short sentence.",
        tools: [{ type: "web_search", mode: "live" }],
      },
      environment: { type: "none" },
      input: [{
        role: "user",
        content: [{
          type: "input_text",
          text: `Search NASA's Mars facts for why Mars looks red. Start the answer with '${token}' and include a NASA source link.`,
        }],
      }],
    },
  });
  const schedulePath = `${automations}/${encodeURIComponent(schedule.id)}`;
  console.log(`Created ${schedule.id}. Waiting for a scheduled delivery.`);

  let invocation;
  try {
    invocation = await waitFor("a scheduled invocation", async () => {
      const { data } = await request(`${schedulePath}/invocations?limit=10`);
      const scheduled = data.find((item) => item.scheduled_at !== null && item.source.type === "automation");
      if (scheduled?.status === "failed" || scheduled?.status === "skipped") {
        throw new Error(`Scheduled delivery ${scheduled.status}: ${scheduled.error?.code ?? scheduled.skip_reason ?? scheduled.id}`);
      }
      if (scheduled?.status === "completed" && scheduled.session_id) return scheduled;
    }, 120_000);
    console.log(`Scheduled invocation ${invocation.id} started session ${invocation.session_id}.`);
  } finally {
    await request(schedulePath, "DELETE");
    console.log(`Removed ${schedule.id}.`);
  }

  await waitFor("the agent's web search and final answer", async () => {
    const { data } = await request(`/v1/agents/sessions/${encodeURIComponent(invocation.session_id)}/items?limit=50`);
    const searched = data.some((item) => item.type === "web_search_call" && item.status === "completed");
    const answered = data.some((item) => item.type === "message" && item.role === "assistant" &&
      item.phase === "final_answer" && item.status === "completed" &&
      item.content.some((part) => part.type === "output_text" && part.text.includes(token)));
    return searched && answered;
  }, 90_000);
  console.log("PASS: The one-minute automation ran web search and returned a final answer.");
}

for (const signal of ["SIGINT", "SIGTERM"]) {
  process.once(signal, () => stop.abort(new Error(`Stopped by ${signal}`)));
}

try {
  if (!apiKey) throw new Error("Set OPENAI_API_KEY in the environment.");
  await request(`${automations}?limit=1`);
  console.log("PASS: The Agents automations endpoint is accessible.");
  if (values.schedule) await verifySchedule();
} catch (error) {
  console.error(error.message);
  process.exitCode = 1;
}
