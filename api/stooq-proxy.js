// Vercel Serverless Function - pobiera dane ze Stooq po stronie serwera
// Dostepna pod: /api/stooq-proxy?s=SYMBOL&d1=YYYYMMDD&d2=YYYYMMDD
export default async function handler(req, res) {
  const { s, d1, d2 } = req.query;

  if (!s) {
    res.status(400).json({ error: "Brak parametru s (symbol)" });
    return;
  }

  let stooqUrl = `https://stooq.pl/q/d/l/?s=${encodeURIComponent(String(s).toLowerCase())}&i=d`;
  if (d1) stooqUrl += `&d1=${d1}`;
  if (d2) stooqUrl += `&d2=${d2}`;

  try {
    const resp = await fetch(stooqUrl, {
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept": "text/csv,text/plain,*/*",
        "Referer": "https://stooq.pl/",
      },
    });
    const text = await resp.text();
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.setHeader("Cache-Control", "public, max-age=60");
    res.status(200).send(text);
  } catch (e) {
    res.status(502).json({ error: String(e.message || e) });
  }
}
