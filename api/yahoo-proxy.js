// Vercel Serverless Function - pobiera dane z Yahoo Finance po stronie serwera
// Dostepna pod: /api/yahoo-proxy?symbol=CDR.WA&range=3mo&interval=1d
export default async function handler(req, res) {
  const { symbol, range, interval } = req.query;

  if (!symbol) {
    res.status(400).json({ error: "Brak parametru symbol" });
    return;
  }

  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}?range=${encodeURIComponent(range || "3mo")}&interval=${encodeURIComponent(interval || "1d")}`;

  try {
    const resp = await fetch(url, {
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept": "application/json",
      },
    });
    const data = await resp.json();
    const result = data?.chart?.result?.[0];
    if (!result) {
      res.status(404).json({ error: "Brak danych dla " + symbol });
      return;
    }
    const timestamps = result.timestamp || [];
    const closes = result.indicators?.quote?.[0]?.close || [];
    const rows = timestamps
      .map((t, i) => ({ date: new Date(t * 1000).toISOString().slice(0, 10), close: closes[i] }))
      .filter((r) => Number.isFinite(r.close) && r.close > 0);

    res.setHeader("Cache-Control", "public, max-age=60");
    res.status(200).json({ rows });
  } catch (e) {
    res.status(502).json({ error: String(e.message || e) });
  }
}
