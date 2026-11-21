import { readFile } from "node:fs/promises";
import { createServer as createHttpServer } from "node:http";
import { createServer as createHttpsServer } from "node:https";
import { extname, resolve, sep } from "node:path";

const types = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".xml": "application/xml; charset=utf-8",
};
const root = resolve("dist");
const key = process.env.SHEETDOT_TLS_KEY;
const cert = process.env.SHEETDOT_TLS_CERT;
if (Boolean(key) !== Boolean(cert)) throw new Error("Set both SHEETDOT_TLS_KEY and SHEETDOT_TLS_CERT");

async function serve(request, response) {
  if (request.method !== "GET" && request.method !== "HEAD") {
    response.writeHead(405, { Allow: "GET, HEAD" }).end();
    return;
  }

  try {
    const url = new URL(request.url, "http://localhost");
    const pathname = decodeURIComponent(url.pathname === "/" ? "/runtime.html" : url.pathname);
    const file = resolve(root, `.${pathname}`);
    if (!file.startsWith(`${root}${sep}`) || !types[extname(file)]) {
      response.writeHead(404).end();
      return;
    }

    const body = await readFile(file);
    response.writeHead(200, {
      "Content-Type": types[extname(file)],
      "Cache-Control": "no-store",
      "Access-Control-Allow-Origin": "*",
    });
    response.end(request.method === "GET" ? body : undefined);
  } catch (error) {
    const status = error instanceof URIError ? 400 : error.code === "ENOENT" ? 404 : 500;
    response.writeHead(status).end();
  }
}

const server = key && cert
  ? createHttpsServer({ key: await readFile(key), cert: await readFile(cert) }, serve)
  : createHttpServer(serve);
const port = Number(process.env.SHEETDOT_PORT ?? 3000);
server.listen(port, "0.0.0.0", () => console.log(`Serving dist on ${key ? "HTTPS" : "HTTP"} port ${port}`));
