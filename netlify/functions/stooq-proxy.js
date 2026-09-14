// Standardowa Netlify Function - pobiera dane ze Stooq po stronie serwera
// Dostepna pod: /.netlify/functions/stooq-proxy?s=SYMBOL&d1=YYYYMMDD&d2=YYYYMMDD
export async function handler(event) {
  const params = event.queryStringParameters || {};
  const s = params.s;
  const d1 = params.d1;
  const d2 = params.d2;

  if (!s) {
    return {
      statusCode: 400,
      body: JSON.stringify({ error: "Brak parametru s (symbol)" }),
    };
  }

  let stooqUrl = `https://stooq.pl/q/d/l/?s=${encodeURIComponent(s.toLowerCase())}&i=d`;
  if (d1) stooqUrl += `&d1=${d1}`;
  if (d2) stooqUrl += `&d2=${d2}`;

  try {
    const resp = await fetch(stooqUrl);
    const text = await resp.text();
    return {
      statusCode: 200,
      headers: {
        "Content-Type": "text/plain; charset=utf-8",
        "Cache-Control": "public, max-age=60",
      },
      body: text,
    };
  } catch (e) {
    return {
      statusCode: 502,
      body: JSON.stringify({ error: String(e.message || e) }),
    };
  }
}
