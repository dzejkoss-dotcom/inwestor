// Vercel Serverless Function - pobiera dane z Yahoo Finance po stronie serwera
// Dostepna pod: /api/yahoo-proxy?symbol=CDR.WA&range=3mo&interval=1d
export default async function handler(req, res) {
  const { symbol, range, interval, events } = req.query;

  if (!symbol) {
    res.status(400).json({ error: "Brak parametru symbol" });
    return;
  }

  let url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}?range=${encodeURIComponent(range || "3mo")}&interval=${encodeURIComponent(interval || "1d")}`;
  if (events) url += `&events=${encodeURIComponent(events)}`;

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
      .map((t, i) => ({ date: new Date(t * 1000).toISOString(), close: closes[i] }))
      .filter((r) => Number.isFinite(r.close) && r.close > 0);

    // Yahoo potrafi przez kilka dni nie aktualizować dziennych notowań (np. GPW).
    // Notowania godzinowe są aktualne, więc ostatnie dni nadpisujemy ostatnią ceną godzinową z każdego dnia.
    if ((interval || "1d") === "1d") {
      try {
        const hUrl = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}?range=1mo&interval=60m`;
        const hResp = await fetch(hUrl, { headers: { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36", "Accept": "application/json" } });
        const h = (await hResp.json())?.chart?.result?.[0];
        const hts = h?.timestamp || [];
        const hcl = h?.indicators?.quote?.[0]?.close || [];
        const lastByDay = {};
        hts.forEach((t, i) => {
          if (Number.isFinite(hcl[i]) && hcl[i] > 0) lastByDay[new Date(t * 1000).toISOString().slice(0, 10)] = { date: new Date(t * 1000).toISOString(), close: hcl[i] };
        });
        const firstHourDay = Object.keys(lastByDay).sort()[0];
        if (firstHourDay) {
          const byDay = {};
          for (const r of rows) byDay[r.date.slice(0, 10)] = r;
          for (const [day, r] of Object.entries(lastByDay)) {
            if (byDay[day]) byDay[day].close = r.close;
            else if (day >= (rows[0]?.date.slice(0, 10) || firstHourDay)) rows.push(r);
          }
          rows.sort((a, b) => a.date.localeCompare(b.date));
        }
      } catch (e) {}
    }

    const dividends = result.events?.dividends
      ? Object.values(result.events.dividends).map((d) => ({
          date: new Date(d.date * 1000).toISOString().slice(0, 10),
          amount: d.amount,
        }))
      : [];

    // Bieżąca cena i poprzednie zamknięcie dokładnie tak, jak pokazuje je Yahoo/Google,
    // plus godziny sesji giełdy (do kropki "rynek otwarty")
    const m = result.meta || {};
    const reg = m.currentTradingPeriod?.regular;
    const meta = {
      price: m.regularMarketPrice ?? null,
      previousClose: m.previousClose ?? m.chartPreviousClose ?? null,
      currency: m.currency ?? null,
      time: m.regularMarketTime ? new Date(m.regularMarketTime * 1000).toISOString() : null,
      sessionStart: reg?.start ? reg.start * 1000 : null,
      sessionEnd: reg?.end ? reg.end * 1000 : null,
      exchange: m.exchangeName ?? null,
    };

    res.setHeader("Cache-Control", "public, max-age=60");
    res.status(200).json({ rows, dividends, meta });
  } catch (e) {
    res.status(502).json({ error: String(e.message || e) });
  }
}
