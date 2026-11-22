const rowsPerFeed = 20;
const githubReleases = "https://api.github.com/repos/nodejs/node/releases?per_page=12";
const housingSource = "https://data.cityofnewyork.us/resource/vy5i-a666.json?" +
  new URLSearchParams({ $where: "development_type = 'Rental'", $order: "lottery_start_date DESC, lottery_id DESC", $limit: "12" });

// Excel's day zero. Local wall-clock dates and UTC timestamps are formatted and labeled separately.
const serial = value => (typeof value === "number" ? value : Date.parse(
  /(?:Z|[+-]\d\d:\d\d)$/.test(value) ? value : value + (value.length === 10 ? "T00:00:00Z" : "Z")
)) / 86_400_000 + 25_569;

const tubeRows = lines => lines.map(line => [line.name,
  line.lineStatuses.map(status => status.statusSeverityDescription).join("; "),
  [...new Set(line.lineStatuses.map(status => status.reason).filter(Boolean))].join(" • "), line.id]);

const feeds = {
  "bitcoin-trades": {
    url: "https://api.exchange.coinbase.com/products/BTC-USD/trades?limit=5",
    headers: ["Time (IST)", "Price (USD)"],
    rows: trades => trades.filter(trade => trade.time && Number(trade.price) > 0)
      .sort((a, b) => b.time.localeCompare(a.time))
      .map(trade => [new Intl.DateTimeFormat("en-GB", {
        timeZone: "Asia/Kolkata", hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23",
      }).format(new Date(trade.time)), Number(trade.price)]),
  },
  "bengaluru-hours": {
    url: "https://api.open-meteo.com/v1/forecast?latitude=12.97&longitude=77.59&hourly=temperature_2m,precipitation_probability&forecast_hours=4&timezone=Asia%2FKolkata",
    headers: ["Time (IST)", "Temperature °C", "Rain chance %"],
    rows: ({ hourly }) => hourly.time.map((time, index) =>
      [time, hourly.temperature_2m[index], hourly.precipitation_probability[index]]),
  },
  weather: {
    url: "https://api.open-meteo.com/v1/forecast?latitude=12.9716&longitude=77.5946&daily=temperature_2m_max,temperature_2m_min,precipitation_probability_max,precipitation_sum&timezone=Asia%2FKolkata&forecast_days=7",
    headers: ["Day (IST)", "Low °C", "High °C", "Max rain chance", "Rain mm"],
    rows: ({ daily: d }) => d.time.map((time, i) => [serial(time), d.temperature_2m_min[i],
      d.temperature_2m_max[i], d.precipitation_probability_max[i] === null ? null : d.precipitation_probability_max[i] / 100,
      d.precipitation_sum[i]]),
  },
  air: {
    url: "https://air-quality-api.open-meteo.com/v1/air-quality?latitude=12.9716&longitude=77.5946&hourly=pm2_5,pm10,us_aqi&timezone=Asia%2FKolkata&forecast_hours=12",
    headers: ["Hour (IST)", "PM2.5 µg/m³", "PM10 µg/m³", "US AQI"],
    rows: ({ hourly: h }) => h.time.map((time, i) => [serial(time), h.pm2_5[i], h.pm10[i], h.us_aqi[i]]),
  },
  tube: {
    url: "https://api.tfl.gov.uk/Line/Mode/tube/Status",
    headers: ["Tube line", "Service status", "TfL reason", "Line ID"],
    rows: tubeRows,
  },
  earthquakes: {
    url: "https://earthquake.usgs.gov/earthquakes/feed/v1.0/summary/4.5_week.geojson",
    headers: ["Location", "Event time (UTC)", "Magnitude", "Depth km", "Review", "USGS event"],
    rows: ({ features }) => features.filter(({ properties: p }) => p.type === "earthquake")
      .sort((a, b) => b.properties.time - a.properties.time).map(({ properties: p, geometry }) =>
        [p.place ?? p.title, serial(p.time), p.mag, geometry.coordinates[2], p.status, p.url]),
  },
  fx: {
    url: "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml",
    headers: ["Currency", "Units per EUR", "Rate date"],
    xml: true,
    rows: xml => {
      const date = xml.match(/\btime=['"](\d{4}-\d{2}-\d{2})['"]/)?.[1];
      if (!date) throw new Error("The ECB feed has no rate date.");
      const selected = new Set(["USD", "INR", "GBP", "JPY", "CAD", "AUD", "CHF", "SGD"]);
      return Array.from(xml.matchAll(/\bcurrency=['"]([A-Z]{3})['"]\s+rate=['"]([\d.]+)['"]/g))
        .filter(([, currency]) => selected.has(currency))
        .map(([, currency, rate]) => [currency, Number(rate), serial(date)]);
    },
  },
  releases: {
    url: githubReleases,
    headers: ["Version", "Published (UTC)", "Channel", "Release notes"],
    rows: releases => releases.filter(release => !release.draft && release.published_at)
      .sort((a, b) => b.published_at.localeCompare(a.published_at))
      .map(release => [release.tag_name, serial(release.published_at),
        release.prerelease || /^v?\d+\.\d+\.\d+-/.test(release.tag_name) ? "Pre-release" : "Stable", release.html_url]),
  },
  housing: {
    url: housingSource,
    headers: ["Development", "Borough", "Published status", "Opens (NY)", "Closes (NY)", "Units", "Lottery ID"],
    rows: lotteries => lotteries.map(row => [row.lottery_name, row.borough, row.lottery_status,
      serial(row.lottery_start_date.slice(0, 10)), serial(row.lottery_end_date.slice(0, 10)),
      row.unit_count === undefined ? null : Number(row.unit_count), row.lottery_id]),
  },
};

export async function fetchDemoFeed(id) {
  if (!Object.hasOwn(feeds, id)) return null;
  const { url, headers, rows, xml } = feeds[id];
  const response = await fetch(url, {
    headers: { "User-Agent": "Sheetdot-Excel-Demo" },
    signal: AbortSignal.timeout(12_000),
  });
  if (!response.ok) throw new Error(`Public source returned ${response.status}.`);
  const data = rows(xml ? await response.text() : await response.json()).slice(0, rowsPerFeed);
  if (!data.length || data.some(row => row.length !== headers.length ||
    row.some(value => value != null && typeof value !== "string" && !Number.isFinite(value)))) {
    throw new Error("The public source returned no usable rows.");
  }
  return { headers, rows: data, retrieved: new Date().toISOString(), url };
}

export async function demoFeedRequest(request) {
  if (request.method !== "GET") return new Response("Use GET.", { status: 405, headers: { Allow: "GET" } });
  const id = new URL(request.url).searchParams.get("id");
  try {
    const feed = await fetchDemoFeed(id);
    if (!feed) return Response.json({ error: "Unknown demo feed." }, { status: 404 });
    return Response.json(feed, { headers: { "Cache-Control": "no-store" } });
  } catch (error) {
    console.warn("Could not refresh public demo feed", id, error.message);
    return Response.json({ error: "The public source could not be read. Try again shortly." }, { status: 502 });
  }
}

export const bengaluruWeatherTool = "get_bengaluru_forecast";
export const bitcoinTradesTool = "get_coinbase_btc_usd_trades";

export function bengaluruWeatherMcp(request) {
  return feedMcp(request, "bengaluru-hours", bengaluruWeatherTool,
    "Read hourly temperature and rain chance for Bengaluru, India from Open-Meteo. All times are IST.");
}

export function bitcoinTradesMcp(request) {
  return feedMcp(request, "bitcoin-trades", bitcoinTradesTool,
    "Read the five newest real BTC-USD trades from Coinbase Exchange, newest first, with times in IST.");
}

async function feedMcp(request, id, toolName, description) {
  if (request.method !== "POST") {
    return new Response(null, { status: 405, headers: { Allow: "POST" } });
  }
  const message = await request.json().catch(() => null);
  const payload = message && message.jsonrpc === "2.0" ? message : null;
  const reply = body => Response.json({ jsonrpc: "2.0", id: payload?.id ?? null, ...body },
    { headers: { "Cache-Control": "no-store" } });
  if (!payload) return reply({ error: { code: -32600, message: "Invalid request" } });
  const server = {
    serverInfo: { name: "sheetdot-" + id, version: "1.0.0" },
    capabilities: { tools: {} },
  };
  switch (payload.method) {
    case "initialize":
      return reply({ result: { ...server, protocolVersion: "2025-11-25" } });
    case "notifications/initialized":
      return new Response(null, { status: 202 });
    case "ping":
      return reply({ result: {} });
    case "tools/list":
      return reply({ result: { tools: [{
        name: toolName,
        description,
        inputSchema: { type: "object", properties: {}, additionalProperties: false },
        annotations: { readOnlyHint: true },
      }] } });
    case "tools/call": {
      if (payload.params?.name !== toolName || Object.keys(payload.params?.arguments ?? {}).length) {
        return reply({ error: { code: -32602, message: "Unknown tool or invalid arguments" } });
      }
      try {
        const feed = await fetchDemoFeed(id);
        return reply({ result: { content: [{ type: "text", text: JSON.stringify(feed) }] } });
      } catch (error) {
        console.warn("Could not read public demo feed", id, error.message);
        return reply({ result: { isError: true, content: [{ type: "text", text: "The live feed is temporarily unavailable." }] } });
      }
    }
    default:
      return reply({ error: { code: -32601, message: "Method not found" } });
  }
}
