import { cp, mkdir, readFile, writeFile } from "node:fs/promises";
import { build } from "esbuild";

const origin = new URL(process.env.SHEETDOT_PUBLIC_URL ?? "https://localhost:3000");
if (origin.protocol !== "https:" || origin.pathname !== "/" || origin.search ||
    origin.hash || origin.username || origin.password) {
  throw new Error("SHEETDOT_PUBLIC_URL must be an HTTPS origin, such as https://localhost:3000");
}

await mkdir("dist", { recursive: true });
await cp("public", "dist", { recursive: true });
await build({
  entryPoints: {
    functions: "src/functions/compatibility.ts",
    runtime: "src/runtime.ts",
    dialog: "src/dialog.ts",
  },
  outdir: "dist",
  format: "iife",
  bundle: true,
  target: "es2020",
});

const manifest = await readFile("manifest.xml", "utf8");
await writeFile("dist/manifest.xml", manifest.replaceAll("https://localhost:3000", origin.origin));
console.log(`Built dist/manifest.xml for ${origin.origin}`);
