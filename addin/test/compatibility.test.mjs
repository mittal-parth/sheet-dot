import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { runInNewContext } from "node:vm";
import { transform } from "esbuild";

const source = await readFile(new URL("../src/functions/compatibility.ts", import.meta.url), "utf8");
const { code } = await transform(source, { loader: "ts" });
const metadata = JSON.parse(await readFile(new URL("../public/functions.json", import.meta.url), "utf8"));

function loadFunctions() {
  const functions = new Map();
  const timers = new Map();
  runInNewContext(code, {
    CustomFunctions: { associate: (id, fn) => functions.set(id, fn) },
    setTimeout: (fn) => {
      const id = Symbol();
      timers.set(id, fn);
      return id;
    },
    clearTimeout: (id) => timers.delete(id),
  });
  return { functions, timers };
}

test("the registered functions match the manifest metadata", () => {
  const { functions } = loadFunctions();
  assert.deepEqual([...functions.keys()], metadata.functions.map(({ id }) => id));
  assert.equal(functions.get("ECHO")("Value returned"), "Value returned");
  const rows = functions.get("SPILL")();
  assert.equal(JSON.stringify(rows), JSON.stringify([["Row 1", "Spill worked"], ["Row 2", "2 columns"]]));
});

test("the streaming probe emits both values and cancels pending work", () => {
  const { functions, timers } = loadFunctions();
  const results = [];
  const first = { setResult: (value) => results.push(value) };
  const second = { setResult: () => {} };
  functions.get("STREAM")("Stream finished", first);
  functions.get("STREAM")("Canceled", second);
  second.onCanceled();

  assert.deepEqual(results, ["Working…"]);
  assert.equal(timers.size, 1);
  for (const finish of timers.values()) finish();
  assert.deepEqual(results, ["Working…", "Stream finished"]);
});
