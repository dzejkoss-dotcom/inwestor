import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

// Poza artefaktami Claude window.storage nie istnieje — ten sam interfejs
// dostajemy tutaj z localStorage przeglądarki, żeby kod działał wszędzie bez zmian.
if (typeof window !== "undefined" && !window.storage) {
  window.storage = {
    async get(key, _shared) {
      const raw = window.localStorage.getItem(key);
      if (raw === null) return null;
      return { key, value: raw, shared: !!_shared };
    },
    async set(key, value, _shared) {
      window.localStorage.setItem(key, value);
      return { key, value, shared: !!_shared };
    },
    async delete(key, _shared) {
      window.localStorage.removeItem(key);
      return { key, deleted: true, shared: !!_shared };
    },
    async list(prefix, _shared) {
      const keys = Object.keys(window.localStorage).filter((k) => !prefix || k.startsWith(prefix));
      return { keys, prefix, shared: !!_shared };
    },
  };
}
import {
  PieChart, Pie, Cell, ResponsiveContainer, AreaChart, Area,
  XAxis, YAxis, Tooltip, CartesianGrid, BarChart, Bar, LabelList,
  ComposedChart, Line, Scatter, Legend,
} from "recharts";
import {
  Plus, TrendingUp, TrendingDown, RefreshCw, X, Wallet,
  Trash2, ChevronRight, Building2, Calendar, Hash, Upload,
  Settings, Pencil, Check, LayoutDashboard, Briefcase, BarChart3,
  ArrowLeft,
} from "lucide-react";

const DEFAULT_PORTFOLIOS = ["ING", "XTB", "IKE", "PZU TFI"];
const SEED_XTB_TRANSACTIONS = [
  { ticker: "TTWO.US", name: "Take-Two Interactive", type: "buy", quantity: 0.0239, price: 868.6192, currency: "PLN", date: "2026-08-06" },
  { ticker: "TTWO.US", name: "Take-Two Interactive", type: "sell", quantity: 0.0239, price: 876.569, currency: "PLN", date: "2026-08-28" },
  { ticker: "TTWO.US", name: "Take-Two Interactive", type: "buy", quantity: 0.9761, price: 867.9131, currency: "PLN", date: "2026-08-06" },
  { ticker: "TTWO.US", name: "Take-Two Interactive", type: "sell", quantity: 0.9761, price: 876.7544, currency: "PLN", date: "2026-08-28" },
  { ticker: "TTWO.US", name: "Take-Two Interactive", type: "buy", quantity: 0.2215, price: 867.9007, currency: "PLN", date: "2026-08-06" },
  { ticker: "TTWO.US", name: "Take-Two Interactive", type: "sell", quantity: 0.2215, price: 877.3815, currency: "PLN", date: "2026-08-28" },
  { ticker: "SPCX.US", name: "SpaceX", type: "buy", quantity: 2, price: 429.995, currency: "PLN", date: "2026-08-05" },
  { ticker: "SPCX.US", name: "SpaceX", type: "sell", quantity: 2, price: 409.765, currency: "PLN", date: "2026-08-06" },
  { ticker: "SPCX.US", name: "SpaceX", type: "buy", quantity: 0.5328, price: 429.9925, currency: "PLN", date: "2026-08-05" },
  { ticker: "SPCX.US", name: "SpaceX", type: "sell", quantity: 0.5328, price: 409.7598, currency: "PLN", date: "2026-08-06" },
  { ticker: "AMZN.US", name: "Amazon", type: "buy", quantity: 1, price: 892.61, currency: "PLN", date: "2026-07-30" },
  { ticker: "AMZN.US", name: "Amazon", type: "sell", quantity: 1, price: 997.04, currency: "PLN", date: "2026-07-31" },
  { ticker: "AMZN.US", name: "Amazon", type: "buy", quantity: 0.0947, price: 892.6082, currency: "PLN", date: "2026-07-30" },
  { ticker: "AMZN.US", name: "Amazon", type: "sell", quantity: 0.0947, price: 995.8817, currency: "PLN", date: "2026-07-31" },
  { ticker: "CCC.PL", name: "Modivo", type: "buy", quantity: 0.7299, price: 78.0244, currency: "PLN", date: "2026-05-08" },
  { ticker: "CCC.PL", name: "Modivo", type: "sell", quantity: 0.7299, price: 92.9853, currency: "PLN", date: "2026-06-23" },
  { ticker: "CCC.PL", name: "Modivo", type: "buy", quantity: 5.2701, price: 78.0213, currency: "PLN", date: "2026-05-08" },
  { ticker: "CCC.PL", name: "Modivo", type: "sell", quantity: 5.2701, price: 92.9793, currency: "PLN", date: "2026-06-23" },
  { ticker: "CCC.PL", name: "Modivo", type: "buy", quantity: 31.2475, price: 78.021, currency: "PLN", date: "2026-05-08" },
  { ticker: "CCC.PL", name: "Modivo", type: "sell", quantity: 31.2475, price: 93.0001, currency: "PLN", date: "2026-06-23" },
  { ticker: "CCC.PL", name: "Modivo", type: "buy", quantity: 32.7525, price: 89.0201, currency: "PLN", date: "2026-04-21" },
  { ticker: "CCC.PL", name: "Modivo", type: "sell", quantity: 32.7525, price: 92.9999, currency: "PLN", date: "2026-06-23" },
  { ticker: "CCC.PL", name: "Modivo", type: "buy", quantity: 23.4146, price: 89.0201, currency: "PLN", date: "2026-04-21" },
  { ticker: "CCC.PL", name: "Modivo", type: "sell", quantity: 23.4146, price: 93.0001, currency: "PLN", date: "2026-06-23" },
  { ticker: "CCC.PL", name: "Modivo", type: "buy", quantity: 0.5854, price: 90.246, currency: "PLN", date: "2026-04-21" },
  { ticker: "CCC.PL", name: "Modivo", type: "sell", quantity: 0.5854, price: 92.9962, currency: "PLN", date: "2026-06-23" },
  { ticker: "CCC.PL", name: "Modivo", type: "buy", quantity: 64, price: 90.2427, currency: "PLN", date: "2026-04-21" },
  { ticker: "CCC.PL", name: "Modivo", type: "sell", quantity: 64, price: 93, currency: "PLN", date: "2026-06-23" },
  { ticker: "CCC.PL", name: "Modivo", type: "buy", quantity: 21, price: 90.2429, currency: "PLN", date: "2026-04-21" },
  { ticker: "CCC.PL", name: "Modivo", type: "sell", quantity: 21, price: 93, currency: "PLN", date: "2026-06-23" },
  { ticker: "CCC.PL", name: "Modivo", type: "buy", quantity: 25, price: 90.2428, currency: "PLN", date: "2026-04-21" },
  { ticker: "CCC.PL", name: "Modivo", type: "sell", quantity: 25, price: 93, currency: "PLN", date: "2026-06-23" },
  { ticker: "CCC.PL", name: "Modivo", type: "buy", quantity: 0.2302, price: 90.2259, currency: "PLN", date: "2026-04-21" },
  { ticker: "CCC.PL", name: "Modivo", type: "sell", quantity: 0.2302, price: 93.0061, currency: "PLN", date: "2026-06-23" },
  { ticker: "XTB.PL", name: "XTB", type: "buy", quantity: 1, price: 107.36, currency: "PLN", date: "2026-04-23" },
  { ticker: "XTB.PL", name: "XTB", type: "sell", quantity: 1, price: 101.12, currency: "PLN", date: "2026-05-08" },
  { ticker: "XTB.PL", name: "XTB", type: "buy", quantity: 27, price: 107.36, currency: "PLN", date: "2026-04-23" },
  { ticker: "XTB.PL", name: "XTB", type: "sell", quantity: 27, price: 101.14, currency: "PLN", date: "2026-05-08" },
  { ticker: "XTB.PL", name: "XTB", type: "buy", quantity: 4, price: 88.94, currency: "PLN", date: "2026-02-27" },
  { ticker: "XTB.PL", name: "XTB", type: "sell", quantity: 4, price: 107, currency: "PLN", date: "2026-04-13" },
  { ticker: "XTB.PL", name: "XTB", type: "buy", quantity: 175, price: 89.3008, currency: "PLN", date: "2026-02-27" },
  { ticker: "XTB.PL", name: "XTB", type: "sell", quantity: 175, price: 107, currency: "PLN", date: "2026-04-13" },
  { ticker: "TSM.US", name: "TSMC", type: "buy", quantity: 3.8396, price: 1600.5964, currency: "PLN", date: "2026-08-28" },
  { ticker: "TSM.US", name: "TSMC", type: "buy", quantity: 0.077, price: 1600.6494, currency: "PLN", date: "2026-08-28" },
];
const CURRENCIES = ["PLN", "USD", "EUR"];
const STORAGE_KEY = "portfolio:transactions";
const PORTFOLIOS_KEY = "portfolio:portfolios";
const PRICES_KEY = "portfolio:prices-cache";
const DIVIDENDS_KEY = "portfolio:dividends-cache";
const TWELVEDATA_KEY_STORAGE = "portfolio:twelvedata-api-key";
const SEED_FLAG_KEY = "portfolio:xtb-seed-2026-09-08-v2-applied";
const SEED_TAG = "xtb-file-seed";
const OLD_SEED_XTB_V1 = [
  "SPCX.US|buy|2|429.995|2026-08-05", "SPCX.US|sell|2|409.765|2026-08-06",
  "SPCX.US|buy|0.5328|429.9925|2026-08-05", "SPCX.US|sell|0.5328|409.7598|2026-08-06",
  "AMZN.US|buy|1|892.61|2026-07-30", "AMZN.US|sell|1|997.04|2026-07-31",
  "AMZN.US|buy|0.0947|892.6082|2026-07-30", "AMZN.US|sell|0.0947|995.8817|2026-07-31",
  "TTWO.US|buy|1.1976|858.4256|2026-08-06", "TTWO.US|buy|0.0239|859.2096|2026-08-06",
];
function isOldSeedV1(tx) {
  return OLD_SEED_XTB_V1.includes(`${tx.ticker}|${tx.type}|${tx.quantity}|${tx.price}|${tx.date}`);
}

const SEED_IKE_FLAG_KEY = "portfolio:ike-seed-2026-09-08-applied";
const SEED_IKE_TAG = "ike-file-seed";
const SEED_IKE_TRANSACTIONS = [
  { ticker: "MSFT.US", name: "Microsoft", type: "buy", quantity: 5, price: 1540.908, currency: "PLN", date: "2026-01-30" },
  { ticker: "MSFT.US", name: "Microsoft", type: "sell", quantity: 5, price: 1554.498, currency: "PLN", date: "2026-06-04" },
  { ticker: "MSFT.US", name: "Microsoft", type: "buy", quantity: 0.1464, price: 1540.9153, currency: "PLN", date: "2026-01-30" },
  { ticker: "MSFT.US", name: "Microsoft", type: "sell", quantity: 0.1464, price: 1554.235, currency: "PLN", date: "2026-06-04" },
  { ticker: "AMZN.US", name: "Amazon", type: "buy", quantity: 8, price: 831.81, currency: "PLN", date: "2025-10-27" },
  { ticker: "AMZN.US", name: "Amazon", type: "sell", quantity: 8, price: 921.2688, currency: "PLN", date: "2026-06-04" },
  { ticker: "AMZN.US", name: "Amazon", type: "buy", quantity: 0.6868, price: 831.8142, currency: "PLN", date: "2025-10-27" },
  { ticker: "AMZN.US", name: "Amazon", type: "sell", quantity: 0.6868, price: 921.1998, currency: "PLN", date: "2026-06-04" },
  { ticker: "ETFBM40TR.PL", name: "mWIG40TR", type: "buy", quantity: 0.0263, price: 128.8973, currency: "PLN", date: "2025-10-27" },
  { ticker: "ETFBM40TR.PL", name: "mWIG40TR", type: "sell", quantity: 0.0263, price: 146.3878, currency: "PLN", date: "2026-05-08" },
  { ticker: "ETFBM40TR.PL", name: "mWIG40TR", type: "buy", quantity: 1.3184, price: 129.1035, currency: "PLN", date: "2025-10-27" },
  { ticker: "ETFBM40TR.PL", name: "mWIG40TR", type: "sell", quantity: 1.3184, price: 146.2985, currency: "PLN", date: "2026-05-08" },
  { ticker: "ETFBM40TR.PL", name: "mWIG40TR", type: "buy", quantity: 8.6553, price: 129.1001, currency: "PLN", date: "2025-10-27" },
  { ticker: "ETFBM40TR.PL", name: "mWIG40TR", type: "sell", quantity: 8.6553, price: 146.3, currency: "PLN", date: "2026-05-08" },
  { ticker: "ETFBM40TR.PL", name: "mWIG40TR", type: "buy", quantity: 7, price: 129.1, currency: "PLN", date: "2025-10-27" },
  { ticker: "ETFBM40TR.PL", name: "mWIG40TR", type: "sell", quantity: 7, price: 146.4, currency: "PLN", date: "2026-05-08" },
  { ticker: "ETFBM40TR.PL", name: "mWIG40TR", type: "buy", quantity: 50, price: 129.1, currency: "PLN", date: "2025-10-27" },
  { ticker: "ETFBM40TR.PL", name: "mWIG40TR", type: "sell", quantity: 50, price: 146.42, currency: "PLN", date: "2026-05-08" },
  { ticker: "ETFBM40TR.PL", name: "mWIG40TR", type: "buy", quantity: 0.2615, price: 129.1013, currency: "PLN", date: "2025-10-27" },
  { ticker: "ETFBM40TR.PL", name: "mWIG40TR", type: "sell", quantity: 0.2615, price: 146.4245, currency: "PLN", date: "2026-05-08" },
  { ticker: "ISAC.UK", name: "MSCI ACWI", type: "buy", quantity: 21, price: 394.841, currency: "PLN", date: "2025-10-27" },
  { ticker: "ISAC.UK", name: "MSCI ACWI", type: "sell", quantity: 21, price: 394.3695, currency: "PLN", date: "2026-01-30" },
  { ticker: "ISAC.UK", name: "MSCI ACWI", type: "buy", quantity: 0.9454, price: 394.817, currency: "PLN", date: "2025-10-27" },
  { ticker: "ISAC.UK", name: "MSCI ACWI", type: "sell", quantity: 0.9454, price: 394.4891, currency: "PLN", date: "2026-01-30" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "buy", quantity: 0.0189, price: 213.7566, currency: "PLN", date: "2026-01-12" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "sell", quantity: 0.0189, price: 205.291, currency: "PLN", date: "2026-01-30" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "buy", quantity: 0.9453, price: 213.4984, currency: "PLN", date: "2026-01-12" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "sell", quantity: 0.9453, price: 205.5009, currency: "PLN", date: "2026-01-30" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "buy", quantity: 8.0358, price: 213.4996, currency: "PLN", date: "2026-01-12" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "sell", quantity: 8.0358, price: 205.5004, currency: "PLN", date: "2026-01-30" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "buy", quantity: 10.9642, price: 213.5003, currency: "PLN", date: "2026-01-12" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "sell", quantity: 10.9642, price: 205.4997, currency: "PLN", date: "2026-01-30" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "buy", quantity: 7.0358, price: 213.4995, currency: "PLN", date: "2026-01-12" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "sell", quantity: 7.0358, price: 205.5004, currency: "PLN", date: "2026-01-30" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "buy", quantity: 7.9642, price: 213.5004, currency: "PLN", date: "2026-01-12" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "sell", quantity: 7.9642, price: 205.4996, currency: "PLN", date: "2026-01-30" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "buy", quantity: 11, price: 213, currency: "PLN", date: "2026-01-12" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "sell", quantity: 11, price: 205.5, currency: "PLN", date: "2026-01-30" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "buy", quantity: 0.0358, price: 213.1285, currency: "PLN", date: "2026-01-12" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "sell", quantity: 0.0358, price: 205.5866, currency: "PLN", date: "2026-01-30" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "buy", quantity: 0.8662, price: 212.9993, currency: "PLN", date: "2026-01-12" },
  { ticker: "CBF.PL", name: "Cyber_Folks", type: "sell", quantity: 0.8662, price: 205.4953, currency: "PLN", date: "2026-01-30" },
  { ticker: "MSFT.US", name: "Microsoft", type: "buy", quantity: 6.4825, price: 1540.9071, currency: "PLN", date: "2026-01-30" },
  { ticker: "MSFT.US", name: "Microsoft", type: "buy", quantity: 0.2331, price: 1541.184, currency: "PLN", date: "2026-01-30" },
  { ticker: "MSFT.US", name: "Microsoft", type: "buy", quantity: 0.0047, price: 1540.4255, currency: "PLN", date: "2026-01-30" },
  { ticker: "AMZN.US", name: "Amazon", type: "buy", quantity: 9.5855, price: 843.8423, currency: "PLN", date: "2026-01-20" },
  { ticker: "AMZN.US", name: "Amazon", type: "buy", quantity: 1.7369, price: 831.8095, currency: "PLN", date: "2025-10-27" },
  { ticker: "AMZN.US", name: "Amazon", type: "buy", quantity: 0.1992, price: 843.8253, currency: "PLN", date: "2026-01-20" },
  { ticker: "DAT.PL", name: "Datawalk", type: "buy", quantity: 74.0096, price: 130.7702, currency: "PLN", date: "2026-05-08" },
  { ticker: "DAT.PL", name: "Datawalk", type: "buy", quantity: 2.5202, price: 138.001, currency: "PLN", date: "2026-08-05" },
  { ticker: "DAT.PL", name: "Datawalk", type: "buy", quantity: 1.2723, price: 130.5981, currency: "PLN", date: "2026-05-08" },
  { ticker: "DAT.PL", name: "Datawalk", type: "buy", quantity: 0.0504, price: 138.0952, currency: "PLN", date: "2026-08-05" },
  { ticker: "CDR.PL", name: "CD Projekt RED", type: "buy", quantity: 35.3793, price: 277.0001, currency: "PLN", date: "2026-01-19" },
  { ticker: "CDR.PL", name: "CD Projekt RED", type: "buy", quantity: 0.7076, price: 277.0068, currency: "PLN", date: "2026-01-19" },
  { ticker: "CDR.PL", name: "CD Projekt RED", type: "buy", quantity: 0.1351, price: 247.0022, currency: "PLN", date: "2026-03-13" },
  { ticker: "CDR.PL", name: "CD Projekt RED", type: "buy", quantity: 0.0141, price: 276.5957, currency: "PLN", date: "2026-01-19" },
  { ticker: "CDR.PL", name: "CD Projekt RED", type: "buy", quantity: 0.0122, price: 269.6721, currency: "PLN", date: "2026-01-20" },
  { ticker: "CRQ.PL", name: "Creotech Quantum", type: "buy", quantity: 39.8884, price: 393.2549, currency: "PLN", date: "2026-06-05" },
];

const PIE_COLORS = ["#fbbf24", "#34d399", "#60a5fa", "#f87171", "#a78bfa", "#fb923c", "#2dd4bf", "#f472b6"];

const CATEGORY_MAP = {
  "SPCX.US": "Technologia",
  "AMZN.US": "Technologia",
  "MSFT.US": "Technologia",
  "TSM.US": "AI i półprzewodniki",
  "TTWO.US": "Gaming",
  "CDR.PL": "Gaming",
  "CCC.PL": "Handel detaliczny",
  "XTB.PL": "Finanse",
  "CBF.PL": "Technologia",
  "DAT.PL": "AI i półprzewodniki",
  "CRQ.PL": "AI i półprzewodniki",
  "ETFBM40TR.PL": "ETF / Indeksy",
  "ISAC.UK": "ETF / Indeksy",
};
const CATEGORY_COLORS = {
  "Technologia": "#60a5fa",
  "AI i półprzewodniki": "#a78bfa",
  "Gaming": "#f472b6",
  "Handel detaliczny": "#fb923c",
  "Finanse": "#34d399",
  "ETF / Indeksy": "#fbbf24",
  "Inne": "#94a3b8",
};

const LOGO_DOMAIN_MAP = {
  "SPCX.US": "spacex.com",
  "AMZN.US": "amazon.com",
  "MSFT.US": "microsoft.com",
  "TSM.US": "tsmc.com",
  "TTWO.US": "take2games.com",
  "CDR.PL": "cdprojekt.com",
  "CCC.PL": "modivo.pl",
  "XTB.PL": "xtb.com",
  "CBF.PL": "cyberfolks.pl",
  "DAT.PL": "datawalk.com",
  "CRQ.PL": "creotech.pl",
  "ISAC.UK": "ishares.com",
};

function CompanyIcon({ ticker, size = 44 }) {
  const [sourceIndex, setSourceIndex] = useState(0);
  const domain = LOGO_DOMAIN_MAP[ticker.toUpperCase()];

  const sources = domain
    ? [
        `https://logo.clearbit.com/${domain}?size=128`,
        `https://www.google.com/s2/favicons?domain=${domain}&sz=128`,
      ]
    : [];

  if (!domain || sourceIndex >= sources.length) {
    const cat = CATEGORY_MAP[ticker.toUpperCase()] || "Inne";
    const color = CATEGORY_COLORS[cat] || "#94a3b8";
    return (
      <div
        className="rounded-xl flex items-center justify-center text-xs font-bold text-slate-950 tabular-nums"
        style={{ width: size, height: size, background: color }}
      >
        {ticker.slice(0, 3)}
      </div>
    );
  }
  return (
    <div className="rounded-xl bg-white flex items-center justify-center overflow-hidden" style={{ width: size, height: size }}>
      <img
        src={sources[sourceIndex]}
        alt=""
        className="w-full h-full object-contain p-1"
        onError={() => setSourceIndex((i) => i + 1)}
      />
    </div>
  );
}

const BENCHMARKS = [
  { key: "wig", label: "WIG", symbol: "WIG.WA" },
  { key: "wig20", label: "WIG20", symbol: "WIG20.WA" },
  { key: "wig40", label: "WIG40 (mWIG40)", symbol: "MWIG40.WA" },
  { key: "sp500", label: "S&P 500", symbol: "^GSPC" },
  { key: "nasdaq100", label: "NASDAQ 100", symbol: "^NDX" },
];

function fmtPLN(n) {
  const v = Number.isFinite(n) ? n : 0;
  return new Intl.NumberFormat("pl-PL", { style: "currency", currency: "PLN", maximumFractionDigits: 2 }).format(v);
}
function fmtPLNShort(n) {
  const v = Number.isFinite(n) ? n : 0;
  if (Math.abs(v) >= 1000000) return (v / 1000000).toFixed(1).replace(/\.0$/, "") + " mln";
  if (Math.abs(v) >= 1000) return Math.round(v / 1000) + "k";
  return Math.round(v).toString();
}
function fmtPct(n) {
  const v = Number.isFinite(n) ? n : 0;
  const sign = v > 0 ? "+" : "";
  return `${sign}${v.toFixed(2)}%`;
}
function getMarketStatus(ticker, liveOpen) {
  const upper = ticker.toUpperCase();
  let tz, openMin, closeMin;
  if (upper.endsWith(".US")) { tz = "America/New_York"; openMin = 9 * 60 + 30; closeMin = 16 * 60; }
  else if (upper.endsWith(".PL")) { tz = "Europe/Warsaw"; openMin = 9 * 60; closeMin = 17 * 60; }
  else if (upper.endsWith(".UK")) { tz = "Europe/London"; openMin = 8 * 60; closeMin = 16 * 60 + 30; }
  else tz = null;

  let isWeekend = false;
  if (tz) {
    try {
      const weekday = new Intl.DateTimeFormat("en-US", { timeZone: tz, weekday: "short" }).format(new Date());
      isWeekend = weekday === "Sat" || weekday === "Sun";
    } catch (e) {}
  }

  if (liveOpen === true) return "green";
  if (liveOpen === false) return isWeekend ? "gray" : "gold";

  if (!tz) return "gray";
  try {
    const parts = new Intl.DateTimeFormat("en-US", { timeZone: tz, weekday: "short", hour: "2-digit", minute: "2-digit", hour12: false }).formatToParts(new Date());
    const get = (type) => parts.find((p) => p.type === type)?.value;
    const weekday = get("weekday");
    if (weekday === "Sat" || weekday === "Sun") return "gray";
    const hour = parseInt(get("hour"), 10);
    const minute = parseInt(get("minute"), 10);
    const nowMin = hour * 60 + minute;
    return nowMin >= openMin && nowMin < closeMin ? "green" : "gold";
  } catch (e) {
    return "gray";
  }
}
function calculateXIRR(cashflows) {
  if (cashflows.length < 2) return 0;
  const t0 = cashflows[0].date;
  const yearsFrom = (d) => (d.getTime() - t0.getTime()) / (365.25 * 24 * 60 * 60 * 1000);
  function npv(rate) {
    return cashflows.reduce((sum, cf) => sum + cf.amount / Math.pow(1 + rate, yearsFrom(cf.date)), 0);
  }
  function npvDerivative(rate) {
    return cashflows.reduce((sum, cf) => {
      const y = yearsFrom(cf.date);
      if (y === 0) return sum;
      return sum - (y * cf.amount) / Math.pow(1 + rate, y + 1);
    }, 0);
  }
  let rate = 0.1;
  for (let i = 0; i < 100; i++) {
    const f = npv(rate);
    const fp = npvDerivative(rate);
    if (Math.abs(fp) < 1e-10) break;
    const next = rate - f / fp;
    if (!Number.isFinite(next) || next <= -0.999) break;
    if (Math.abs(next - rate) < 1e-7) { rate = next; break; }
    rate = next;
  }
  return Number.isFinite(rate) ? rate * 100 : 0;
}

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}
function txFingerprint(t) {
  const round = (n) => (Number.isFinite(Number(n)) ? Number(n).toFixed(6) : "0");
  return [t.portfolioId, t.ticker.toUpperCase().trim(), t.type, round(t.quantity), round(t.price), t.date].join("|");
}
function parseNum(v) {
  if (v == null) return NaN;
  if (typeof v === "number") return v;
  const cleaned = String(v).replace(/\s|\u00A0/g, "").replace(",", ".");
  return parseFloat(cleaned);
}
function normalizeDate(v) {
  if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0, 10);
  if (typeof v === "number" && v > 20000 && v < 80000) {
    const ms = Math.round((v - 25569) * 86400 * 1000);
    const d = new Date(ms);
    if (!isNaN(d)) return d.toISOString().slice(0, 10);
  }
  if (typeof v === "string") {
    const s = v.trim();
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    m = s.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})/);
    if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
    const num = Number(s);
    if (Number.isFinite(num) && num > 20000 && num < 80000) return normalizeDate(num);
  }
  return "";
}

function useCountUp(target, duration = 700, resetKey) {
  const [val, setVal] = useState(target);
  const prevValueRef = useRef(target);
  const prevResetRef = useRef(resetKey);
  const rafRef = useRef(null);
  useEffect(() => {
    const resetTriggered = resetKey !== undefined && resetKey !== prevResetRef.current;
    prevResetRef.current = resetKey;
    const from = resetTriggered ? 0 : prevValueRef.current;
    const to = target;
    if (!Number.isFinite(from) || !Number.isFinite(to)) { setVal(to); return; }
    const start = performance.now();
    cancelAnimationFrame(rafRef.current);
    function tick(now) {
      const p = Math.min(1, (now - start) / duration);
      const eased = 1 - Math.pow(1 - p, 3);
      setVal(from + (to - from) * eased);
      if (p < 1) rafRef.current = requestAnimationFrame(tick);
      else prevValueRef.current = to;
    }
    rafRef.current = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(rafRef.current);
  }, [target, duration, resetKey]);
  return val;
}

function computePortfolioData(transactions) {
  const groups = {};
  for (const t of transactions) {
    const key = `${t.portfolioId}|${t.ticker.toUpperCase().trim()}`;
    if (!groups[key]) groups[key] = [];
    groups[key].push(t);
  }

  const perPortfolioHoldings = [];
  let realizedPnL = 0;
  let totalInvested = 0;
  const realizedByPortfolio = {};
  const investedByPortfolio = {};

  for (const key in groups) {
    const txs = [...groups[key]].sort((a, b) => new Date(a.date) - new Date(b.date));
    const ticker = txs[0].ticker.toUpperCase().trim();
    const portfolioId = txs[0].portfolioId;
    let name = ticker;
    let currency = txs[0].currency || "PLN";
    let qty = 0, costBasis = 0;
    for (const t of txs) {
      if (t.name?.trim()) name = t.name.trim();
      const q = Number(t.quantity) || 0;
      const p = Number(t.price) || 0;
      if (t.type === "buy") {
        qty += q;
        costBasis += q * p;
        totalInvested += q * p;
        investedByPortfolio[portfolioId] = (investedByPortfolio[portfolioId] || 0) + q * p;
      } else {
        const avg = qty > 0 ? costBasis / qty : 0;
        const sellQty = Math.min(q, qty);
        realizedPnL += (p - avg) * sellQty;
        realizedByPortfolio[portfolioId] = (realizedByPortfolio[portfolioId] || 0) + (p - avg) * sellQty;
        qty -= q;
        costBasis -= avg * sellQty;
        if (qty < 0) qty = 0;
        if (costBasis < 0) costBasis = 0;
      }
    }
    if (qty > 0.00001) {
      perPortfolioHoldings.push({ portfolioId, ticker, name, qty, costBasis, avgCost: costBasis / qty, currency });
    }
  }

  const mergedMap = {};
  for (const h of perPortfolioHoldings) {
    if (!mergedMap[h.ticker]) {
      mergedMap[h.ticker] = { ticker: h.ticker, name: h.name, qty: 0, costBasis: 0, currency: h.currency, portfolioIds: [] };
    }
    const m = mergedMap[h.ticker];
    m.qty += h.qty;
    m.costBasis += h.costBasis;
    m.portfolioIds.push(h.portfolioId);
    if (h.name) m.name = h.name;
  }
  const holdings = Object.values(mergedMap).map((h) => ({ ...h, avgCost: h.qty ? h.costBasis / h.qty : 0 }));

  return { holdings, perPortfolioHoldings, realizedPnL, totalInvested, realizedByPortfolio, investedByPortfolio };
}

async function fetchTwelveDataBatch(symbols, exchange, apiKey) {
  if (!symbols.length) return {};
  let url = `https://api.twelvedata.com/quote?symbol=${encodeURIComponent(symbols.join(","))}&apikey=${encodeURIComponent(apiKey)}`;
  if (exchange) url += `&exchange=${encodeURIComponent(exchange)}`;
  const resp = await fetch(url);
  const data = await resp.json();
  if (data.status === "error" || data.code >= 400) throw new Error(data.message || "Błąd Twelve Data");
  return data.symbol ? { [data.symbol]: data } : data;
}

async function fetchTwelveDataIntraday(symbol, exchange, apiKey) {
  let url = `https://api.twelvedata.com/time_series?symbol=${encodeURIComponent(symbol)}&interval=5min&outputsize=100&apikey=${encodeURIComponent(apiKey)}`;
  if (exchange) url += `&exchange=${encodeURIComponent(exchange)}`;
  const resp = await fetch(url);
  const data = await resp.json();
  if (data.status === "error" || !Array.isArray(data.values)) throw new Error(data.message || "Brak danych śróddziennych");
  return data.values
    .map((v) => ({ t: new Date(v.datetime.replace(" ", "T")).getTime(), close: parseFloat(v.close) }))
    .filter((v) => Number.isFinite(v.close))
    .sort((a, b) => a.t - b.t);
}

function formatStooqDate(d) {
  return d.toISOString().slice(0, 10).replace(/-/g, "");
}

async function fetchStooqSeries(symbol, d1, d2) {
  const end = d2 || new Date();
  const start = d1 || new Date(end.getTime() - 14 * 24 * 60 * 60 * 1000);
  const query = `s=${encodeURIComponent(symbol.toLowerCase())}&d1=${formatStooqDate(start)}&d2=${formatStooqDate(end)}`;
  const proxyUrl = `/api/stooq-proxy?${query}`;
  const directUrl = `https://stooq.pl/q/d/l/?${query}&i=d`;

  let text = null;
  try {
    const resp = await fetch(proxyUrl);
    if (resp.ok) text = await resp.text();
  } catch (e) {}

  if (text === null) {
    const resp = await fetch(directUrl);
    if (!resp.ok) throw new Error("Błąd sieci Stooq");
    text = await resp.text();
  }

  if (/nie znaleziono|brak danych/i.test(text)) throw new Error("Stooq: brak danych dla " + symbol);
  const lines = text.trim().split("\n").filter(Boolean);
  if (lines.length < 2) throw new Error("Stooq: brak danych dla " + symbol);
  const rows = lines.slice(1).map((line) => {
    const [date, , , , close] = line.split(",");
    return { date, close: parseFloat(close) };
  }).filter((r) => Number.isFinite(r.close) && r.close > 0);
  if (!rows.length) throw new Error("Stooq: brak poprawnych cen dla " + symbol);
  return rows;
}

function toYahooSymbol(ticker) {
  const upper = ticker.toUpperCase();
  if (upper.endsWith(".PL")) return upper.replace(/\.PL$/, "") + ".WA";
  if (upper.endsWith(".US")) return upper.replace(/\.US$/, "");
  if (upper.endsWith(".UK")) return upper.replace(/\.UK$/, "") + ".L";
  return upper;
}

async function fetchYahooSeries(yahooSymbol, range, interval) {
  const url = `/api/yahoo-proxy?symbol=${encodeURIComponent(yahooSymbol)}&range=${range || "3mo"}&interval=${interval || "1d"}`;
  const resp = await fetch(url);
  const data = await resp.json();
  if (!resp.ok || data.error) throw new Error(data.error || "Błąd Yahoo Finance");
  if (!data.rows || !data.rows.length) throw new Error("Yahoo: brak danych dla " + yahooSymbol);
  return data.rows;
}

async function fetchYahooDividends(yahooSymbol) {
  const url = `/api/yahoo-proxy?symbol=${encodeURIComponent(yahooSymbol)}&range=max&interval=1mo&events=div`;
  const resp = await fetch(url);
  const data = await resp.json();
  if (!resp.ok || data.error) throw new Error(data.error || "Błąd Yahoo Finance");
  return data.dividends || [];
}

function quantityHeldOnDate(tickerTxSorted, dateMs) {
  let qty = 0;
  for (const t of tickerTxSorted) {
    if (new Date(t.date).getTime() > dateMs) break;
    qty += t.type === "buy" ? Number(t.quantity) : -Number(t.quantity);
  }
  return qty;
}

async function fetchYahooQuote(ticker, usdPlnRate) {
  const upper = ticker.toUpperCase();
  const isUS = upper.endsWith(".US");
  const isPL = upper.endsWith(".PL");
  const yahooSymbol = toYahooSymbol(ticker);
  const rows = await fetchYahooSeries(yahooSymbol, "5d");
  const last = rows[rows.length - 1];
  const prev = rows.length >= 2 ? rows[rows.length - 2] : null;
  let changeAbsolute = 0, changePercent = 0;
  if (prev) {
    changeAbsolute = last.close - prev.close;
    changePercent = (changeAbsolute / prev.close) * 100;
  }
  if (isPL || !isUS) {
    return { price: last.close, changeAbsolute, changePercent, currency: "PLN", asOf: last.date };
  }
  const fx = usdPlnRate ?? (await fetchYahooSeries("PLN=X", "5d")).slice(-1)[0].close;
  return { price: last.close * fx, changeAbsolute: changeAbsolute * fx, changePercent, currency: "PLN", asOf: last.date };
}

async function fetchStooqQuote(ticker, usdPlnRate) {
  const upper = ticker.toUpperCase();
  const isUS = upper.endsWith(".US");
  const isPL = upper.endsWith(".PL");
  if (!isUS && !isPL) throw new Error("Stooq: rynek nieobsługiwany dla " + ticker);
  const stooqSymbol = isPL ? ticker.replace(/\.PL$/i, "") : ticker;
  const rows = await fetchStooqSeries(stooqSymbol);
  const last = rows[rows.length - 1];
  const prev = rows.length >= 2 ? rows[rows.length - 2] : null;
  let changeAbsolute = 0, changePercent = 0;
  if (prev) {
    changeAbsolute = last.close - prev.close;
    changePercent = (changeAbsolute / prev.close) * 100;
  }
  if (isPL) {
    return { price: last.close, changeAbsolute, changePercent, currency: "PLN", asOf: last.date };
  }
  const fx = usdPlnRate ?? (await fetchStooqSeries("usdpln")).slice(-1)[0].close;
  return { price: last.close * fx, changeAbsolute: changeAbsolute * fx, changePercent, currency: "PLN", asOf: last.date };
}

async function fetchTickerQuote(ticker) {
  const prompt = `Wyszukaj aktualny kurs akcji spółki o tickerze/symbolu "${ticker}", najprawdopodobniej notowanej na Giełdzie Papierów Wartościowych w Warszawie (GPW). Jeśli to nie polska spółka, znajdź właściwy rynek. Odpowiedz WYŁĄCZNIE obiektem JSON, bez żadnego innego tekstu, w dokładnie takiej strukturze: {"price": number, "currency": "PLN", "changePercent": number, "changeAbsolute": number, "asOf": "krótki opis daty/godziny danych"}. Jeśli nie znajdziesz wiarygodnych danych, zwróć {"price": null}.`;
  const resp = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1000,
      messages: [{ role: "user", content: prompt }],
      tools: [{ type: "web_search_20250305", name: "web_search" }],
    }),
  });
  const data = await resp.json();
  const text = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
  const clean = text.replace(/```json|```/g, "").trim();
  const match = clean.match(/\{[\s\S]*\}/);
  if (!match) throw new Error("Brak danych w odpowiedzi");
  const parsed = JSON.parse(match[0]);
  if (parsed.price == null) throw new Error("Nie znaleziono kursu");
  return parsed;
}

export default function App() {
  const [transactions, setTransactions] = useState([]);
  const [portfolios, setPortfolios] = useState([]);
  const [prices, setPrices] = useState({});
  const [dividends, setDividends] = useState({ total: 0, byYear: {}, lastUpdated: null, error: "" });
  const [loaded, setLoaded] = useState(false);
  const [refreshing, setRefreshing] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [showHistory, setShowHistory] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [activeTab, setActiveTab] = useState("dashboard");
  const [selectedTicker, setSelectedTicker] = useState(null);
  const [filterPortfolio, setFilterPortfolio] = useState("all");
  const [allocMode, setAllocMode] = useState("amount");
  const [profitMode, setProfitMode] = useState("percent");
  const [refreshError, setRefreshError] = useState("");
  const [twelveDataKey, setTwelveDataKey] = useState("");

  useEffect(() => {
    (async () => {
      let txRaw = [];
      let pfList = null;
      try {
        const t = await window.storage.get(STORAGE_KEY, false);
        if (t?.value) txRaw = JSON.parse(t.value);
      } catch (e) {}
      try {
        const pf = await window.storage.get(PORTFOLIOS_KEY, false);
        if (pf?.value) pfList = JSON.parse(pf.value);
      } catch (e) {}
      if (!pfList) pfList = DEFAULT_PORTFOLIOS.map((name) => ({ id: uid(), name }));

      const nameToId = {};
      pfList.forEach((p) => { nameToId[p.name] = p.id; });

      let seedApplied = false;
      try {
        const flag = await window.storage.get(SEED_FLAG_KEY, false);
        seedApplied = !!flag?.value;
      } catch (e) {}

      if (!seedApplied) {
        let xtbId = nameToId["XTB"];
        if (!xtbId) {
          xtbId = uid();
          pfList.push({ id: xtbId, name: "XTB" });
          nameToId["XTB"] = xtbId;
        }
        txRaw = txRaw.filter((tx) => tx.seedTag !== SEED_TAG && !isOldSeedV1(tx));
        txRaw = [...txRaw, ...SEED_XTB_TRANSACTIONS.map((t) => ({ ...t, portfolioId: xtbId, seedTag: SEED_TAG }))];
        window.storage.set(SEED_FLAG_KEY, "1", false).catch(() => {});
      }

      let ikeSeedApplied = false;
      try {
        const flag2 = await window.storage.get(SEED_IKE_FLAG_KEY, false);
        ikeSeedApplied = !!flag2?.value;
      } catch (e) {}

      if (!ikeSeedApplied) {
        let ikeId = nameToId["IKE"];
        if (!ikeId) {
          ikeId = uid();
          pfList.push({ id: ikeId, name: "IKE" });
          nameToId["IKE"] = ikeId;
        }
        txRaw = txRaw.filter((tx) => tx.seedTag !== SEED_IKE_TAG);
        txRaw = [...txRaw, ...SEED_IKE_TRANSACTIONS.map((t) => ({ ...t, portfolioId: ikeId, seedTag: SEED_IKE_TAG }))];
        window.storage.set(SEED_IKE_FLAG_KEY, "1", false).catch(() => {});
      }

      const migrated = txRaw.map((tx) => {
        if (tx.portfolioId) return tx;
        let pid = nameToId[tx.broker];
        if (!pid) {
          pid = uid();
          const newName = tx.broker || "Inne";
          pfList.push({ id: pid, name: newName });
          nameToId[newName] = pid;
        }
        const { broker, ...rest } = tx;
        return { ...rest, portfolioId: pid };
      });

      try {
        const p = await window.storage.get(PRICES_KEY, false);
        if (p?.value) setPrices(JSON.parse(p.value));
      } catch (e) {}
      try {
        const d = await window.storage.get(DIVIDENDS_KEY, false);
        if (d?.value) setDividends(JSON.parse(d.value));
      } catch (e) {}
      try {
        const k = await window.storage.get(TWELVEDATA_KEY_STORAGE, false);
        if (k?.value) setTwelveDataKey(k.value);
      } catch (e) {}

      setPortfolios(pfList);
      setTransactions(migrated);
      setLoaded(true);
    })();
  }, []);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set(STORAGE_KEY, JSON.stringify(transactions), false).catch(() => {});
  }, [transactions, loaded]);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set(PORTFOLIOS_KEY, JSON.stringify(portfolios), false).catch(() => {});
  }, [portfolios, loaded]);

  const portfolioName = useCallback((id) => portfolios.find((p) => p.id === id)?.name || "Inne", [portfolios]);

  const { holdings, perPortfolioHoldings, realizedPnL, totalInvested, realizedByPortfolio, investedByPortfolio } = useMemo(() => computePortfolioData(transactions), [transactions]);

  const visibleHoldings = useMemo(() => {
    if (filterPortfolio === "all") return holdings;
    return perPortfolioHoldings.filter((h) => h.portfolioId === filterPortfolio);
  }, [holdings, perPortfolioHoldings, filterPortfolio]);

  const metrics = useMemo(() => {
    let costBasisTotal = 0, currentValueTotal = 0, dailyChangePLN = 0;
    for (const h of holdings) {
      const price = prices[h.ticker]?.price ?? h.avgCost;
      const change = prices[h.ticker]?.changeAbsolute ?? 0;
      costBasisTotal += h.costBasis;
      currentValueTotal += h.qty * price;
      dailyChangePLN += h.qty * change;
    }
    const prevValue = currentValueTotal - dailyChangePLN;
    const dailyChangePercent = prevValue > 0 ? (dailyChangePLN / prevValue) * 100 : 0;
    let years = 0;
    if (transactions.length) {
      const earliest = new Date(Math.min(...transactions.map((t) => new Date(t.date).getTime())));
      years = Math.max((Date.now() - earliest.getTime()) / (1000 * 60 * 60 * 24 * 365.25), 1 / 365.25);
    }
    let cagr = 0;
    if (transactions.length && currentValueTotal > 0) {
      const sorted = [...transactions].sort((a, b) => new Date(a.date) - new Date(b.date));
      const cashflows = sorted.map((t) => ({
        date: new Date(t.date),
        amount: t.type === "buy" ? -(Number(t.quantity) * Number(t.price)) : Number(t.quantity) * Number(t.price),
      }));
      cashflows.push({ date: new Date(), amount: currentValueTotal });
      cagr = calculateXIRR(cashflows);
    }
    const unrealizedPnL = currentValueTotal - costBasisTotal;
    const totalProfit = unrealizedPnL + realizedPnL + dividends.total;
    const totalProfitPercent = totalInvested > 0 ? (totalProfit / totalInvested) * 100 : 0;
    return { costBasisTotal, currentValueTotal, dailyChangePLN, dailyChangePercent, cagr, years, unrealizedPnL, totalProfit, totalProfitPercent };
  }, [holdings, prices, transactions, realizedPnL, totalInvested, dividends.total]);

  const investedOverTime = useMemo(() => {
    const sorted = [...transactions].sort((a, b) => new Date(a.date) - new Date(b.date));
    let running = 0;
    const points = [];
    for (const t of sorted) {
      const q = Number(t.quantity) || 0, p = Number(t.price) || 0;
      running += t.type === "buy" ? q * p : -q * p;
      points.push({ date: t.date, value: Math.max(running, 0) });
    }
    return points;
  }, [transactions]);

  const pieData = useMemo(() => {
    return visibleHoldings.map((h) => ({
      name: h.ticker,
      value: h.qty * (prices[h.ticker]?.price ?? h.avgCost),
    })).filter((d) => d.value > 0);
  }, [visibleHoldings, prices]);

  const [dashVisit, setDashVisit] = useState(0);
  useEffect(() => {
    if (activeTab === "dashboard") setDashVisit((v) => v + 1);
  }, [activeTab]);
  useEffect(() => {
    if (loaded) setDashVisit((v) => v + 1);
  }, [loaded]);

  const animatedValue = useCountUp(metrics.currentValueTotal, 700, dashVisit);
  const animatedDailyChange = useCountUp(metrics.dailyChangePLN, 700, dashVisit);
  const animatedCagr = useCountUp(metrics.cagr, 700, dashVisit);
  const animatedTotalProfit = useCountUp(metrics.totalProfit, 700, dashVisit);

  const addTransactions = useCallback((txs) => {
    let skippedDuplicates = 0;
    setTransactions((prev) => {
      const existingFingerprints = new Set(prev.map(txFingerprint));
      const toAdd = [];
      for (const t of txs) {
        const fp = txFingerprint(t);
        if (existingFingerprints.has(fp)) {
          skippedDuplicates++;
          continue;
        }
        existingFingerprints.add(fp);
        toAdd.push({ ...t, id: uid() });
      }
      return [...prev, ...toAdd];
    });
    return skippedDuplicates;
  }, []);

  const deleteTransaction = useCallback((id) => {
    setTransactions((prev) => prev.filter((t) => t.id !== id));
  }, []);

  const addPortfolio = useCallback((name) => {
    const p = { id: uid(), name };
    setPortfolios((prev) => [...prev, p]);
    return p;
  }, []);

  const renamePortfolio = useCallback((id, name) => {
    setPortfolios((prev) => prev.map((p) => (p.id === id ? { ...p, name } : p)));
  }, []);

  const deletePortfolio = useCallback((id) => {
    setPortfolios((prev) => prev.filter((p) => p.id !== id));
    setTransactions((prev) => prev.filter((t) => t.portfolioId !== id));
    setFilterPortfolio((cur) => (cur === id ? "all" : cur));
  }, []);

  const refreshPrices = useCallback(async () => {
    if (!holdings.length) return;
    setRefreshing(true);
    setRefreshError("");

    let usdPlnRate = null;
    try {
      const rows = await fetchYahooSeries("PLN=X", "5d");
      usdPlnRate = rows[rows.length - 1].close;
    } catch (e) {}

    const resolved = {};
    let tdCount = 0;
    let tdErrorMsg = "";

    if (twelveDataKey) {
      const usTickers = holdings.filter((h) => h.ticker.toUpperCase().endsWith(".US")).map((h) => h.ticker.replace(/\.US$/i, ""));
      const plTickers = holdings.filter((h) => h.ticker.toUpperCase().endsWith(".PL")).map((h) => h.ticker.replace(/\.PL$/i, ""));
      try {
        const [usData, plData, fxData] = await Promise.all([
          fetchTwelveDataBatch(usTickers, null, twelveDataKey).catch((e) => { tdErrorMsg = tdErrorMsg || `USD: ${String(e.message || e)}`; return {}; }),
          fetchTwelveDataBatch(plTickers, "XWAR", twelveDataKey).catch((e) => { tdErrorMsg = tdErrorMsg || `PLN: ${String(e.message || e)}`; return {}; }),
          usTickers.length ? fetchTwelveDataBatch(["USD/PLN"], null, twelveDataKey).catch((e) => { tdErrorMsg = tdErrorMsg || `FX: ${String(e.message || e)}`; return {}; }) : Promise.resolve({}),
        ]);
        const tdUsdPlnEntry = fxData["USD/PLN"];
        const tdUsdPlnRate = tdUsdPlnEntry && Number.isFinite(parseFloat(tdUsdPlnEntry.close)) ? parseFloat(tdUsdPlnEntry.close) : (usdPlnRate || null);
        for (const h of holdings) {
          const isUS = h.ticker.toUpperCase().endsWith(".US");
          const isPL = h.ticker.toUpperCase().endsWith(".PL");
          const bareSymbol = h.ticker.replace(/\.(US|PL)$/i, "");
          const entry = isUS ? usData[bareSymbol] : isPL ? plData[bareSymbol] : null;
          if (!entry || !entry.close) {
            if (entry && (entry.message || entry.code)) {
              tdErrorMsg = tdErrorMsg || `${bareSymbol}: ${entry.message || ("kod " + entry.code)}`;
            } else if (!tdErrorMsg) {
              const src = isUS ? usData : plData;
              const keys = Object.keys(src);
              tdErrorMsg = `${bareSymbol}: brak w odpowiedzi (dostępne: ${keys.join(", ") || "nic"})`;
            }
            continue;
          }
          const closeNative = parseFloat(entry.close);
          const changeNative = parseFloat(entry.change ?? 0);
          const changePercent = parseFloat(entry.percent_change ?? 0);
          const isMarketOpen = typeof entry.is_market_open === "boolean" ? entry.is_market_open : null;
          if (!Number.isFinite(closeNative)) continue;
          if (isUS) {
            if (!tdUsdPlnRate) { tdErrorMsg = tdErrorMsg || "Brak kursu USD/PLN"; continue; }
            resolved[h.ticker] = { price: closeNative * tdUsdPlnRate, changeAbsolute: changeNative * tdUsdPlnRate, changePercent, currency: "PLN", asOf: entry.datetime || "", source: "twelvedata", isMarketOpen, fetchedAt: Date.now() };
          } else {
            resolved[h.ticker] = { price: closeNative, changeAbsolute: changeNative, changePercent, currency: "PLN", asOf: entry.datetime || "", source: "twelvedata", isMarketOpen, fetchedAt: Date.now() };
          }
          tdCount++;
        }
      } catch (e) { tdErrorMsg = tdErrorMsg || String(e.message || e); }
    }

    const remaining = holdings.filter((h) => !resolved[h.ticker]);
    let yahooErrorMsg = "";
    const results = await Promise.all(
      remaining.map(async (h) => {
        try {
          const q = await fetchYahooQuote(h.ticker, usdPlnRate);
          return { ticker: h.ticker, quote: { ...q, source: "yahoo", fetchedAt: Date.now() } };
        } catch (e) {
          if (!yahooErrorMsg) yahooErrorMsg = `${h.ticker}: ${String(e.message || e)}`;
        }
        try {
          const q = await fetchTickerQuote(h.ticker);
          return { ticker: h.ticker, quote: { ...q, source: "ai", fetchedAt: Date.now() } };
        } catch (e) {
          return { ticker: h.ticker, quote: null };
        }
      })
    );

    const next = { ...prices, ...resolved };
    let yahooCount = 0, aiCount = 0;
    const failedTickers = [];
    for (const r of results) {
      if (r.quote) {
        next[r.ticker] = r.quote;
        if (r.quote.source === "yahoo") yahooCount++;
        else aiCount++;
      } else {
        failedTickers.push(r.ticker);
      }
    }
    setPrices(next);
    window.storage.set(PRICES_KEY, JSON.stringify(next), false).catch(() => {});
    const parts = [];
    if (twelveDataKey) parts.push(`Twelve Data: ${tdCount}/${holdings.length}${tdCount === 0 && tdErrorMsg ? ` (${tdErrorMsg})` : ""}`);
    parts.push(`Yahoo: ${yahooCount}/${remaining.length}${yahooCount === 0 && yahooErrorMsg ? ` (${yahooErrorMsg})` : ""}`);
    parts.push(`AI: ${aiCount}/${remaining.length}`);
    if (failedTickers.length) parts.push(`bez danych: ${failedTickers.join(", ")}`);
    setRefreshError(parts.join(" · "));
    setRefreshing(false);
  }, [holdings, prices, twelveDataKey]);

  const refreshDividends = useCallback(async () => {
    const byTicker = {};
    for (const t of transactions) {
      const key = t.ticker.toUpperCase();
      if (!byTicker[key]) byTicker[key] = [];
      byTicker[key].push(t);
    }
    const tickers = Object.keys(byTicker);
    if (!tickers.length) return;

    let usdPlnRate = null;
    try {
      const rows = await fetchYahooSeries("PLN=X", "5d");
      usdPlnRate = rows[rows.length - 1].close;
    } catch (e) {}

    let total = 0;
    const byYear = {};
    let anyError = "";
    let anySuccess = false;
    for (const ticker of tickers) {
      const sortedTxs = [...byTicker[ticker]].sort((a, b) => new Date(a.date) - new Date(b.date));
      try {
        const yahooSymbol = toYahooSymbol(ticker);
        const divs = await fetchYahooDividends(yahooSymbol);
        const isUS = ticker.toUpperCase().endsWith(".US");
        const fx = isUS ? (usdPlnRate || 1) : 1;
        for (const d of divs) {
          const dateMs = new Date(d.date).getTime();
          const qty = quantityHeldOnDate(sortedTxs, dateMs);
          if (qty > 0) {
            const amountPLN = qty * d.amount * fx;
            total += amountPLN;
            const year = new Date(d.date).getFullYear();
            byYear[year] = (byYear[year] || 0) + amountPLN;
          }
        }
        anySuccess = true;
      } catch (e) {
        anyError = anyError || `${ticker}: ${String(e.message || e)}`;
      }
    }
    const result = { total, byYear, lastUpdated: Date.now(), error: anySuccess ? "" : anyError };
    setDividends(result);
    window.storage.set(DIVIDENDS_KEY, JSON.stringify(result), false).catch(() => {});
  }, [transactions]);

  const isPositive = metrics.dailyChangePLN >= 0;

  if (selectedTicker) {
    return (
      <StockDetailScreen
        ticker={selectedTicker}
        transactions={transactions}
        portfolios={portfolios}
        prices={prices}
        twelveDataKey={twelveDataKey}
        onBack={() => setSelectedTicker(null)}
      />
    );
  }

  return (
    <div className="min-h-full w-full bg-slate-950 text-slate-100 font-sans">
      <style>{`.recharts-wrapper *:focus { outline: none !important; }`}</style>
      <div className="pb-24">
      {activeTab === "dashboard" && (
      <div className="max-w-md mx-auto px-4 pt-6">

        <div className="flex items-center justify-between mb-4">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 rounded-lg bg-amber-400 flex items-center justify-center">
              <Wallet size={16} className="text-slate-900" />
            </div>
            <h1 className="text-lg font-bold tracking-tight">Mój portfel</h1>
          </div>
          <button
            onClick={() => { refreshPrices(); refreshDividends(); }}
            disabled={refreshing || !holdings.length}
            className="flex items-center gap-1.5 text-xs text-slate-400 disabled:opacity-40 active:scale-95 transition-transform"
          >
            <RefreshCw size={14} className={refreshing ? "animate-spin" : ""} />
            {refreshing ? "Odświeżam" : "Odśwież ceny"}
          </button>
        </div>

        <div className="rounded-2xl bg-slate-900 border border-slate-800 p-5 mb-4 relative overflow-hidden">
          <p className="text-xs text-slate-400 mb-1">Wartość portfela</p>
          <p className="text-3xl tabular-nums tracking-tight font-bold tracking-tight mb-3">{fmtPLN(animatedValue)}</p>
          <div className="flex items-center gap-2">
            <div className={`flex items-center gap-1 px-2 py-1 rounded-lg text-xs tabular-nums tracking-tight ${isPositive ? "bg-emerald-950 text-emerald-400" : "bg-rose-950 text-rose-400"}`}>
              {isPositive ? <TrendingUp size={13} /> : <TrendingDown size={13} />}
              {fmtPLN(animatedDailyChange)}
            </div>
            <div className={`px-2 py-1 rounded-lg text-xs tabular-nums tracking-tight ${isPositive ? "bg-emerald-950 text-emerald-400" : "bg-rose-950 text-rose-400"}`}>
              {fmtPct(metrics.dailyChangePercent)}
            </div>
            <span className="text-xs text-slate-500 ml-auto">dzisiaj</span>
          </div>
        </div>

        <div className="grid grid-cols-2 gap-3 mb-5">
          <div className="rounded-xl bg-slate-900 border border-slate-800 p-3.5">
            <p className="text-xs text-slate-400 mb-1">Wynik portfela</p>
            <p className={`text-lg tabular-nums tracking-tight font-bold ${metrics.cagr >= 0 ? "text-emerald-400" : "text-rose-400"}`}>
              {fmtPct(animatedCagr)}
            </p>
          </div>
          <div className="rounded-xl bg-slate-900 border border-slate-800 p-3.5">
            <p className="text-xs text-slate-400 mb-1">Zysk</p>
            <p className={`text-lg tabular-nums tracking-tight font-bold ${metrics.totalProfit >= 0 ? "text-emerald-400" : "text-rose-400"}`}>
              {fmtPLN(animatedTotalProfit)}
            </p>
            <p className={`text-xs tabular-nums tracking-tight ${metrics.totalProfit >= 0 ? "text-emerald-500" : "text-rose-500"}`}>
              {fmtPct(metrics.totalProfitPercent)}
            </p>
            {dividends.total > 0 && (
              <p className="text-[10px] text-slate-500 mt-0.5">w tym dywidendy: {fmtPLN(dividends.total)}</p>
            )}
          </div>
        </div>

        <div className="flex items-center gap-2 mb-4">
          <div className="flex gap-2 overflow-x-auto no-scrollbar flex-1">
            <button
              onClick={() => setFilterPortfolio("all")}
              className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-medium transition-colors ${filterPortfolio === "all" ? "bg-amber-400 text-slate-900" : "bg-slate-900 text-slate-400 border border-slate-800"}`}
            >
              Wszystkie
            </button>
            {portfolios.map((p) => (
              <button
                key={p.id}
                onClick={() => setFilterPortfolio(p.id)}
                className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-medium transition-colors ${filterPortfolio === p.id ? "bg-amber-400 text-slate-900" : "bg-slate-900 text-slate-400 border border-slate-800"}`}
              >
                {p.name}
              </button>
            ))}
          </div>
          <button onClick={() => setActiveTab("ustawienia")} className="shrink-0 w-8 h-8 rounded-full bg-slate-900 border border-slate-800 flex items-center justify-center text-slate-400">
            <Settings size={14} />
          </button>
        </div>

        {visibleHoldings.length === 0 ? (
          <div className="rounded-2xl border border-dashed border-slate-800 p-8 text-center mb-6">
            <p className="text-sm text-slate-400 mb-3">
              {holdings.length === 0 ? "Nie masz jeszcze żadnych transakcji." : "Brak spółek dla tego filtra."}
            </p>
          </div>
        ) : (
          <>
            <div className="rounded-2xl bg-slate-900 border border-slate-800 p-4 mb-6">
              <p className="text-xs text-slate-400 mb-2">Alokacja portfela</p>
              <div style={{ height: 200 }}>
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart key={dashVisit}>
                    <Pie
                      data={pieData}
                      dataKey="value"
                      nameKey="name"
                      innerRadius={45}
                      outerRadius={70}
                      paddingAngle={2}
                      stroke="none"
                      isAnimationActive
                      animationDuration={700}
                      label={({ percent }) => `${(percent * 100).toFixed(0)}%`}
                      labelLine={false}
                    >
                      {pieData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                    </Pie>
                    <Tooltip
                      formatter={(v) => fmtPLN(v)}
                      contentStyle={{ background: "#0f172a", border: "1px solid #1e293b", borderRadius: 8, fontSize: 12 }}
                      itemStyle={{ color: "#e2e8f0" }}
                      labelStyle={{ color: "#94a3b8" }}
                    />
                  </PieChart>
                </ResponsiveContainer>
              </div>
              <div className="flex flex-wrap gap-x-3 gap-y-1.5 mt-2 justify-center">
                {pieData.map((d, i) => {
                  const total = pieData.reduce((s, x) => s + x.value, 0);
                  const pct = total > 0 ? (d.value / total) * 100 : 0;
                  return (
                    <div key={d.name} className="flex items-center gap-1.5 text-xs text-slate-400">
                      <span className="w-2 h-2 rounded-full" style={{ background: PIE_COLORS[i % PIE_COLORS.length] }} />
                      {d.name} · {pct.toFixed(0)}%
                    </div>
                  );
                })}
              </div>
            </div>

            <div className="mb-6">
              <p className="text-xs text-slate-400 mb-2 px-1">Spółki</p>
              <div className="rounded-2xl bg-slate-900 border border-slate-800 divide-y divide-slate-800 overflow-hidden">
                {visibleHoldings.map((h) => {
                  const priceInfo = prices[h.ticker];
                  const price = priceInfo?.price ?? h.avgCost;
                  const value = h.qty * price;
                  const changePct = priceInfo?.changePercent;
                  const changeAbs = priceInfo?.changeAbsolute;
                  const hasChange = changePct != null;
                  const posChange = (changePct ?? 0) >= 0;
                  const status = getMarketStatus(h.ticker, priceInfo?.isMarketOpen);
                  const statusColor = status === "green" ? "#34d399" : status === "gold" ? "#fbbf24" : "#64748b";
                  const typeLabel = CATEGORY_MAP[h.ticker] === "ETF / Indeksy" ? "ETF" : "Akcje";
                  const posPnl = h.qty * (price - h.avgCost);
                  const posPnlPercent = h.avgCost > 0 ? ((price - h.avgCost) / h.avgCost) * 100 : 0;
                  const posPnlPositive = posPnl >= 0;
                  return (
                    <div key={h.ticker} onClick={() => setSelectedTicker(h.ticker)} className="p-3.5 flex items-center gap-3 active:bg-slate-800/50 cursor-pointer">
                      <div className="relative w-11 h-11 shrink-0">
                        <CompanyIcon ticker={h.ticker} size={44} />
                        <span
                          className="absolute -bottom-0.5 -right-0.5 w-3.5 h-3.5 rounded-full border-2 border-slate-900"
                          style={{ background: statusColor }}
                        />
                      </div>
                      <div className="min-w-0 flex-1">
                        <div className="flex items-center gap-1.5">
                          <p className="text-sm font-bold truncate">{h.name}</p>
                          <span className="shrink-0 text-[10px] px-1.5 py-0.5 rounded-full bg-slate-800 text-slate-400">{typeLabel}</span>
                        </div>
                        <p className={`text-xs tabular-nums tracking-tight ${posPnlPositive ? "text-emerald-400" : "text-rose-400"}`}>
                          {fmtPLN(posPnl)} ({fmtPct(posPnlPercent)})
                        </p>
                      </div>
                      <div className="text-right shrink-0">
                        <p className="text-sm tabular-nums tracking-tight font-bold">{fmtPLN(value)}</p>
                        <p className={`text-xs tabular-nums tracking-tight ${!hasChange ? "text-slate-600" : posChange ? "text-emerald-400" : "text-rose-400"}`}>
                          {!hasChange ? "brak kursu" : `${fmtPLN(changeAbs)} (${fmtPct(changePct)})`}
                        </p>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            <div className="rounded-2xl bg-slate-900 border border-slate-800 p-4 mb-6">
              <p className="text-base font-bold mb-4">Największe pozycje</p>
              <div className="flex items-center justify-between mb-3 px-1">
                <span className="text-xs text-slate-500">Instrument</span>
                <div className="flex items-center gap-5">
                  <div className="flex flex-col items-center gap-1">
                    <span className="text-xs text-slate-500">Alokacja</span>
                    <div className="flex bg-slate-800 rounded-full p-0.5">
                      <button
                        onClick={() => setAllocMode("percent")}
                        className={`px-2.5 py-1 rounded-full text-xs font-bold transition-colors ${allocMode === "percent" ? "bg-slate-950 text-slate-100" : "text-slate-500"}`}
                      >
                        %
                      </button>
                      <button
                        onClick={() => setAllocMode("amount")}
                        className={`px-2.5 py-1 rounded-full text-xs font-bold transition-colors ${allocMode === "amount" ? "bg-slate-950 text-slate-100" : "text-slate-500"}`}
                      >
                        zł
                      </button>
                    </div>
                  </div>
                  <div className="flex flex-col items-center gap-1">
                    <span className="text-xs text-slate-500">Zysk</span>
                    <div className="flex bg-slate-800 rounded-full p-0.5">
                      <button
                        onClick={() => setProfitMode("percent")}
                        className={`px-2.5 py-1 rounded-full text-xs font-bold transition-colors ${profitMode === "percent" ? "bg-slate-950 text-slate-100" : "text-slate-500"}`}
                      >
                        %
                      </button>
                      <button
                        onClick={() => setProfitMode("amount")}
                        className={`px-2.5 py-1 rounded-full text-xs font-bold transition-colors ${profitMode === "amount" ? "bg-slate-950 text-slate-100" : "text-slate-500"}`}
                      >
                        zł
                      </button>
                    </div>
                  </div>
                </div>
              </div>

              <div className="space-y-2">
                {[...visibleHoldings]
                  .map((h) => {
                    const priceInfo = prices[h.ticker];
                    const price = priceInfo?.price ?? h.avgCost;
                    const value = h.qty * price;
                    const pnl = h.qty * (price - h.avgCost);
                    const pnlPercent = h.avgCost > 0 ? ((price - h.avgCost) / h.avgCost) * 100 : 0;
                    return { ...h, value, pnl, pnlPercent };
                  })
                  .sort((a, b) => b.value - a.value)
                  .map((h) => {
                    const totalValue = visibleHoldings.reduce((s, x) => {
                      const p = prices[x.ticker]?.price ?? x.avgCost;
                      return s + x.qty * p;
                    }, 0);
                    const maxValue = Math.max(...visibleHoldings.map((x) => x.qty * (prices[x.ticker]?.price ?? x.avgCost)));
                    const allocPercent = totalValue > 0 ? (h.value / totalValue) * 100 : 0;
                    const barPct = maxValue > 0 ? (h.value / maxValue) * 48 : 0;
                    const pnlPositive = h.pnl >= 0;
                    return (
                      <div key={h.ticker} onClick={() => setSelectedTicker(h.ticker)} className="relative rounded-xl overflow-hidden cursor-pointer">
                        <div
                          className="absolute inset-y-0 left-0 bg-indigo-900/50 rounded-xl"
                          style={{ width: `${barPct}%` }}
                        />
                        <div className="relative flex items-center gap-2.5 px-3 py-2.5">
                          <CompanyIcon ticker={h.ticker} size={30} />
                          <p className="text-sm font-bold truncate flex-1">{h.name}</p>
                          <div className="flex items-center gap-2 shrink-0">
                            <p className="text-sm font-bold tabular-nums tracking-tight text-right w-16">
                              {allocMode === "percent" ? `${allocPercent.toFixed(0)}%` : fmtPLN(h.value)}
                            </p>
                            <p className={`text-sm font-bold tabular-nums tracking-tight text-right w-16 ${pnlPositive ? "text-emerald-400" : "text-rose-400"}`}>
                              {profitMode === "percent" ? fmtPct(h.pnlPercent) : fmtPLN(h.pnl)}
                            </p>
                          </div>
                        </div>
                      </div>
                    );
                  })}
              </div>
            </div>

            {Object.keys(dividends.byYear).length > 0 && (
              <div className="rounded-2xl bg-slate-900 border border-slate-800 p-4 mt-4">
                <p className="text-sm font-bold mb-1">Dywidendy wg roku</p>
                <p className="text-xs text-slate-500 mb-3">Łącznie: {fmtPLN(dividends.total)}</p>
                <div style={{ height: 140 }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={Object.entries(dividends.byYear).sort(([a], [b]) => a - b).map(([year, amount]) => ({ year, amount }))}>
                      <XAxis dataKey="year" tick={{ fontSize: 10, fill: "#94a3b8" }} axisLine={{ stroke: "#1e293b" }} tickLine={false} />
                      <YAxis tick={{ fontSize: 10, fill: "#64748b" }} tickFormatter={(v) => fmtPLNShort(v)} width={40} />
                      <Tooltip
                        cursor={{ fill: "transparent" }}
                        formatter={(v) => fmtPLN(v)}
                        contentStyle={{ background: "#0f172a", border: "1px solid #1e293b", borderRadius: 8, fontSize: 12 }}
                        itemStyle={{ color: "#e2e8f0" }}
                        labelStyle={{ color: "#94a3b8" }}
                      />
                      <Bar dataKey="amount" fill="#34d399" radius={[6, 6, 0, 0]} isAnimationActive />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
                {dividends.lastUpdated && (
                  <p className="text-[10px] text-slate-600 mt-2">
                    Zaktualizowano: {new Date(dividends.lastUpdated).toLocaleDateString("pl-PL")}
                  </p>
                )}
              </div>
            )}
            {dividends.error && Object.keys(dividends.byYear).length === 0 && (
              <p className="text-xs text-slate-600 mt-3">Dywidendy: {dividends.error}</p>
            )}
          </>
        )}

      </div>
      )}

      {activeTab === "portfele" && (
        <PortfeleTab
          portfolios={portfolios}
          perPortfolioHoldings={perPortfolioHoldings}
          realizedByPortfolio={realizedByPortfolio}
          investedByPortfolio={investedByPortfolio}
          prices={prices}
          onSelectPortfolio={(id) => { setFilterPortfolio(id); setActiveTab("dashboard"); }}
        />
      )}

      {activeTab === "statystyki" && (
        <StatystykiTab
          transactions={transactions}
          portfolios={portfolios}
          prices={prices}
        />
      )}

      {activeTab === "ustawienia" && (
        <SettingsTab
          portfolios={portfolios}
          transactions={transactions}
          onAdd={addPortfolio}
          onRename={renamePortfolio}
          onDelete={deletePortfolio}
          twelveDataKey={twelveDataKey}
          onSaveTwelveDataKey={(key) => {
            setTwelveDataKey(key);
            window.storage.set(TWELVEDATA_KEY_STORAGE, key, false).catch(() => {});
          }}
          refreshStatus={refreshError}
          onImport={() => setShowImport(true)}
          onAddManual={() => setShowForm(true)}
          onShowHistory={() => setShowHistory(true)}
        />
      )}
      </div>

      <BottomNav activeTab={activeTab} setActiveTab={setActiveTab} />

      {showForm && (
        <TransactionForm
          portfolios={portfolios}
          onManage={() => { setShowForm(false); setActiveTab("ustawienia"); }}
          onClose={() => setShowForm(false)}
          onSubmit={(tx) => { addTransactions([tx]); setShowForm(false); }}
        />
      )}
      {showHistory && (
        <HistorySheet
          transactions={[...transactions].sort((a, b) => new Date(b.date) - new Date(a.date))}
          portfolioName={portfolioName}
          onClose={() => setShowHistory(false)}
          onDelete={deleteTransaction}
        />
      )}
      {showImport && (
        <ImportModal
          portfolios={portfolios}
          onAddPortfolio={addPortfolio}
          onClose={() => setShowImport(false)}
          onImport={(txs) => addTransactions(txs)}
        />
      )}
    </div>
  );
}

function BottomNav({ activeTab, setActiveTab }) {
  const items = [
    { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { key: "portfele", label: "Portfele", icon: Briefcase },
    { key: "statystyki", label: "Statystyki", icon: BarChart3 },
    { key: "ustawienia", label: "Ustawienia", icon: Settings },
  ];
  return (
    <div className="fixed bottom-0 left-0 right-0 bg-slate-900 border-t border-slate-800 z-40">
      <div className="max-w-md mx-auto grid grid-cols-4">
        {items.map((it) => {
          const Icon = it.icon;
          const active = activeTab === it.key;
          return (
            <button
              key={it.key}
              onClick={() => setActiveTab(it.key)}
              className="flex flex-col items-center justify-center gap-1 py-2.5"
            >
              <Icon size={20} className={active ? "text-amber-400" : "text-slate-500"} />
              <span className={`text-[10px] ${active ? "text-amber-400" : "text-slate-500"}`}>{it.label}</span>
            </button>
          );
        })}
      </div>
    </div>
  );
}

function PortfeleTab({ portfolios, perPortfolioHoldings, realizedByPortfolio, investedByPortfolio, prices, onSelectPortfolio }) {
  const rows = portfolios.map((p) => {
    const holdings = perPortfolioHoldings.filter((h) => h.portfolioId === p.id);
    let currentValue = 0, costBasis = 0, dailyChange = 0;
    for (const h of holdings) {
      const info = prices[h.ticker];
      const price = info?.price ?? h.avgCost;
      currentValue += h.qty * price;
      costBasis += h.costBasis;
      dailyChange += h.qty * (info?.changeAbsolute ?? 0);
    }
    const unrealized = currentValue - costBasis;
    const realized = realizedByPortfolio[p.id] || 0;
    const totalPnl = unrealized + realized;
    const invested = investedByPortfolio[p.id] || 0;
    const totalPnlPercent = invested > 0 ? (totalPnl / invested) * 100 : 0;
    const prevValue = currentValue - dailyChange;
    const dailyChangePercent = prevValue > 0 ? (dailyChange / prevValue) * 100 : 0;
    return { ...p, currentValue, dailyChange, dailyChangePercent, totalPnl, totalPnlPercent, holdingsCount: holdings.length };
  });

  return (
    <div className="max-w-md mx-auto px-4 pt-6">
      <h1 className="text-lg font-bold tracking-tight mb-6">Portfele</h1>

      {rows.length === 0 ? (
        <div className="rounded-2xl border border-dashed border-slate-800 p-8 text-center">
          <p className="text-sm text-slate-500">Nie masz jeszcze żadnych portfeli.</p>
        </div>
      ) : (
        <div className="space-y-3">
          {rows.map((p) => {
            const dailyPositive = p.dailyChange >= 0;
            const totalPositive = p.totalPnl >= 0;
            return (
              <button
                key={p.id}
                onClick={() => onSelectPortfolio(p.id)}
                className="w-full text-left rounded-2xl bg-slate-900 border border-slate-800 p-4 active:scale-[0.98] transition-transform"
              >
                <div className="flex items-center justify-between mb-2">
                  <p className="text-base font-bold">{p.name}</p>
                  <ChevronRight size={16} className="text-slate-500" />
                </div>
                <p className="text-2xl tabular-nums tracking-tight font-bold tracking-tight mb-2">{fmtPLN(p.currentValue)}</p>
                <div className="flex items-center gap-2 mb-3">
                  <span className={`px-2 py-1 rounded-lg text-xs tabular-nums tracking-tight ${dailyPositive ? "bg-emerald-950 text-emerald-400" : "bg-rose-950 text-rose-400"}`}>
                    {fmtPLN(p.dailyChange)} ({fmtPct(p.dailyChangePercent)})
                  </span>
                  <span className="text-xs text-slate-500">dzisiaj</span>
                </div>
                <div className="flex items-center justify-between pt-3 border-t border-slate-800">
                  <span className="text-xs text-slate-500">{p.holdingsCount} {p.holdingsCount === 1 ? "spółka" : "spółek"}</span>
                  <span className={`text-sm font-bold tabular-nums ${totalPositive ? "text-emerald-400" : "text-rose-400"}`}>
                    Zysk: {fmtPLN(p.totalPnl)} ({fmtPct(p.totalPnlPercent)})
                  </span>
                </div>
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

function EmptyTab({ icon: Icon, title, text }) {
  return (
    <div className="max-w-md mx-auto px-4 pt-6">
      <h1 className="text-lg font-bold tracking-tight mb-6">{title}</h1>
      <div className="rounded-2xl border border-dashed border-slate-800 p-10 text-center">
        <Icon size={28} className="text-slate-600 mx-auto mb-3" />
        <p className="text-sm text-slate-500">{text}</p>
      </div>
    </div>
  );
}

function SettingsTab({ portfolios, transactions, onAdd, onRename, onDelete, twelveDataKey, onSaveTwelveDataKey, refreshStatus, onImport, onAddManual, onShowHistory }) {
  const [newName, setNewName] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editValue, setEditValue] = useState("");
  const [confirmingId, setConfirmingId] = useState(null);
  const [keyDraft, setKeyDraft] = useState(twelveDataKey || "");
  const [keySaved, setKeySaved] = useState(false);

  function countFor(id) {
    return transactions.filter((t) => t.portfolioId === id).length;
  }

  return (
    <div className="max-w-md mx-auto px-4 pt-6">
      <h1 className="text-lg font-bold tracking-tight mb-6">Ustawienia</h1>

      <p className="text-sm font-medium mb-3">Dane</p>
      <div className="grid grid-cols-2 gap-2 mb-3">
        <button
          onClick={onImport}
          className="flex items-center justify-center gap-1.5 py-2.5 rounded-xl bg-slate-900 border border-slate-800 text-sm text-slate-200 active:scale-[0.98] transition-transform"
        >
          <Upload size={15} /> Importuj plik
        </button>
        <button
          onClick={onAddManual}
          className="flex items-center justify-center gap-1.5 py-2.5 rounded-xl bg-slate-900 border border-slate-800 text-sm text-slate-200 active:scale-[0.98] transition-transform"
        >
          <Plus size={15} /> Dodaj ręcznie
        </button>
      </div>
      <button
        onClick={onShowHistory}
        className="w-full flex items-center justify-between rounded-xl bg-slate-900 border border-slate-800 p-3 mb-8 active:scale-[0.99] transition-transform"
      >
        <span className="text-sm text-slate-300">Historia transakcji ({transactions.length})</span>
        <ChevronRight size={16} className="text-slate-500" />
      </button>

      <p className="text-sm font-medium mb-3">Twoje portfele</p>
      <div className="space-y-2 mb-4">
        {portfolios.map((p) => (
          <div key={p.id} className="flex items-center gap-2 bg-slate-900 border border-slate-800 rounded-xl p-3">
            {editingId === p.id ? (
              <>
                <input
                  value={editValue}
                  onChange={(e) => setEditValue(e.target.value)}
                  className="flex-1 bg-slate-800 border border-slate-700 rounded-lg px-2.5 py-1.5 text-sm text-slate-100"
                  autoFocus
                />
                <button
                  onClick={() => { if (editValue.trim()) onRename(p.id, editValue.trim()); setEditingId(null); }}
                  className="w-8 h-8 rounded-lg bg-emerald-500 text-slate-950 flex items-center justify-center shrink-0"
                >
                  <Check size={14} />
                </button>
              </>
            ) : (
              <>
                <div className="min-w-0 flex-1">
                  <p className="text-sm font-medium truncate">{p.name}</p>
                  <p className="text-xs text-slate-500">{countFor(p.id)} transakcji</p>
                </div>
                {confirmingId === p.id ? (
                  <div className="flex items-center gap-1.5 shrink-0">
                    <span className="text-xs text-slate-400 mr-1">Usunąć wraz z transakcjami?</span>
                    <button onClick={() => { onDelete(p.id); setConfirmingId(null); }} className="px-2 py-1 rounded-lg bg-rose-500 text-slate-950 text-xs font-medium">Tak</button>
                    <button onClick={() => setConfirmingId(null)} className="px-2 py-1 rounded-lg bg-slate-800 text-slate-300 text-xs">Nie</button>
                  </div>
                ) : (
                  <div className="flex items-center gap-1 shrink-0">
                    <button onClick={() => { setEditingId(p.id); setEditValue(p.name); }} className="w-8 h-8 rounded-lg bg-slate-800 text-slate-400 flex items-center justify-center">
                      <Pencil size={13} />
                    </button>
                    <button onClick={() => setConfirmingId(p.id)} className="w-8 h-8 rounded-lg bg-slate-800 text-slate-400 flex items-center justify-center">
                      <Trash2 size={13} />
                    </button>
                  </div>
                )}
              </>
            )}
          </div>
        ))}
      </div>

      <div className="flex gap-2 mb-8">
        <input
          value={newName}
          onChange={(e) => setNewName(e.target.value)}
          placeholder="Nazwa nowego portfela"
          className="flex-1 bg-slate-900 border border-slate-800 rounded-xl px-3 py-2.5 text-sm text-slate-100 placeholder-slate-600"
        />
        <button
          onClick={() => { if (newName.trim()) { onAdd(newName.trim()); setNewName(""); } }}
          className="px-4 rounded-xl bg-amber-400 text-slate-900 font-medium text-sm"
        >
          Dodaj
        </button>
      </div>

      <div className="border-t border-slate-800 pt-4 mb-8">
        <p className="text-sm font-medium mb-1">Klucz API Twelve Data</p>
        <p className="text-xs text-slate-500 mb-3">
          Opcjonalnie — z kluczem ceny są szybsze i pewniejsze. Darmowy klucz (bez karty): twelvedata.com/register
        </p>
        <div className="flex gap-2">
          <input
            value={keyDraft}
            onChange={(e) => { setKeyDraft(e.target.value); setKeySaved(false); }}
            placeholder="Wklej klucz API"
            className="flex-1 bg-slate-900 border border-slate-800 rounded-xl px-3 py-2.5 text-sm text-slate-100 placeholder-slate-600 tabular-nums tracking-tight"
          />
          <button
            onClick={() => { onSaveTwelveDataKey(keyDraft.trim()); setKeySaved(true); }}
            className="px-4 rounded-xl bg-slate-800 text-slate-200 font-medium text-sm"
          >
            {keySaved ? "Zapisano" : "Zapisz"}
          </button>
        </div>
      </div>

      {refreshStatus && (
        <div className="border-t border-slate-800 pt-4">
          <p className="text-sm font-medium mb-2">Status ostatniego odświeżenia cen</p>
          <div className="text-xs text-slate-500 bg-slate-900 border border-slate-800 rounded-lg px-3 py-2">
            {refreshStatus}
          </div>
        </div>
      )}
    </div>
  );
}

function StatystykiTab({ transactions, portfolios, prices }) {
  const [filterPortfolio, setFilterPortfolio] = useState("all");
  const [selectedBenchmarks, setSelectedBenchmarks] = useState([]);
  const [benchmarkResults, setBenchmarkResults] = useState({});
  const [benchmarkSeries, setBenchmarkSeries] = useState([]);
  const [loadingBenchmarks, setLoadingBenchmarks] = useState(false);
  const [benchmarkError, setBenchmarkError] = useState("");

  const filteredTransactions = useMemo(() => {
    if (filterPortfolio === "all") return transactions;
    return transactions.filter((t) => t.portfolioId === filterPortfolio);
  }, [transactions, filterPortfolio]);

  const { holdings, realizedPnL, totalInvested } = useMemo(
    () => computePortfolioData(filteredTransactions),
    [filteredTransactions]
  );

  const localMetrics = useMemo(() => {
    let costBasisTotal = 0, currentValueTotal = 0;
    for (const h of holdings) {
      const price = prices[h.ticker]?.price ?? h.avgCost;
      costBasisTotal += h.costBasis;
      currentValueTotal += h.qty * price;
    }
    const unrealizedPnL = currentValueTotal - costBasisTotal;
    const totalProfit = unrealizedPnL + realizedPnL;
    const totalProfitPercent = totalInvested > 0 ? (totalProfit / totalInvested) * 100 : 0;
    return { totalProfitPercent, currentValueTotal, realizedPnL };
  }, [holdings, prices, realizedPnL, totalInvested]);

  const investedOverTime = useMemo(() => {
    const sorted = [...filteredTransactions].sort((a, b) => new Date(a.date) - new Date(b.date));
    let running = 0;
    const points = [];
    for (const t of sorted) {
      const q = Number(t.quantity) || 0, p = Number(t.price) || 0;
      running += t.type === "buy" ? q * p : -q * p;
      points.push({ date: t.date, value: Math.max(running, 0) });
    }
    return points;
  }, [filteredTransactions]);

  const categoryData = useMemo(() => {
    const map = {};
    for (const h of holdings) {
      const cat = CATEGORY_MAP[h.ticker] || "Inne";
      const price = prices[h.ticker]?.price ?? h.avgCost;
      if (!map[cat]) map[cat] = { name: cat, value: 0, companies: [] };
      map[cat].value += h.qty * price;
      map[cat].companies.push(h.name || h.ticker);
    }
    const list = Object.values(map).sort((a, b) => b.value - a.value);
    const total = list.reduce((s, x) => s + x.value, 0);
    return list.map((d) => ({ ...d, percent: total > 0 ? (d.value / total) * 100 : 0 }));
  }, [holdings, prices]);

  function toggleBenchmark(key) {
    setSelectedBenchmarks((prev) => (prev.includes(key) ? prev.filter((k) => k !== key) : [...prev, key]));
  }

  function priceOnOrBefore(rows, dateMs) {
    let ans = rows[0]?.close;
    for (const r of rows) {
      if (new Date(r.date).getTime() <= dateMs) ans = r.close;
      else break;
    }
    return ans;
  }

  function simulateBenchmarkSeries(sortedTxs, benchRows, nowMs) {
    let units = 0;
    const points = [];
    for (const t of sortedTxs) {
      const dateMs = new Date(t.date).getTime();
      const cash = (t.type === "buy" ? 1 : -1) * Number(t.quantity) * Number(t.price);
      const price = priceOnOrBefore(benchRows, dateMs) || 1;
      units += cash / price;
      points.push(Math.max(units * price, 0));
    }
    const lastPrice = benchRows[benchRows.length - 1].close;
    points.push(Math.max(units * lastPrice, 0));
    return points;
  }

  async function runComparison() {
    if (!selectedBenchmarks.length || !filteredTransactions.length) return;
    setLoadingBenchmarks(true);
    setBenchmarkError("");
    const sortedTxs = [...filteredTransactions].sort((a, b) => new Date(a.date) - new Date(b.date));
    const now = new Date();

    const portfolioSeries = investedOverTime.map((p) => p.value);
    portfolioSeries.push(localMetrics.currentValueTotal);

    const series = {};
    let anyError = "";
    for (const key of selectedBenchmarks) {
      const bench = BENCHMARKS.find((b) => b.key === key);
      try {
        const rows = await fetchYahooSeries(bench.symbol, "5y");
        series[key] = simulateBenchmarkSeries(sortedTxs, rows, now.getTime());
      } catch (e) {
        anyError = anyError || `${bench.label}: ${String(e.message || e)}`;
      }
    }

    const dates = [...sortedTxs.map((t) => t.date), now.toISOString()];
    const merged = dates.map((date, i) => {
      const point = { date, "Twój portfel": portfolioSeries[i] };
      for (const key of selectedBenchmarks) {
        if (series[key]) point[BENCHMARKS.find((b) => b.key === key).label] = series[key][i];
      }
      return point;
    });

    setBenchmarkSeries(merged);
    setBenchmarkResults(series);
    if (Object.keys(series).length === 0 && anyError) setBenchmarkError(anyError);
    setLoadingBenchmarks(false);
  }

  const seriesColors = ["#fbbf24", "#60a5fa", "#34d399", "#f472b6", "#a78bfa", "#fb923c"];
  const noHighlightCursor = { fill: "transparent" };

  return (
    <div className="max-w-md mx-auto px-4 pt-6">
      <h1 className="text-lg font-bold tracking-tight mb-4">Statystyki</h1>

      <div className="flex gap-2 overflow-x-auto no-scrollbar mb-6">
        <button
          onClick={() => setFilterPortfolio("all")}
          className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-medium transition-colors ${filterPortfolio === "all" ? "bg-amber-400 text-slate-900" : "bg-slate-900 text-slate-400 border border-slate-800"}`}
        >
          Wszystkie
        </button>
        {portfolios.map((p) => (
          <button
            key={p.id}
            onClick={() => setFilterPortfolio(p.id)}
            className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-medium transition-colors ${filterPortfolio === p.id ? "bg-amber-400 text-slate-900" : "bg-slate-900 text-slate-400 border border-slate-800"}`}
          >
            {p.name}
          </button>
        ))}
      </div>

      {investedOverTime.length > 1 ? (
        <div className="rounded-2xl bg-slate-900 border border-slate-800 p-4 mb-6">
          <p className="text-xs text-slate-400 mb-3">Zainwestowany kapitał w czasie</p>
          <div style={{ height: 160 }}>
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={investedOverTime}>
                <defs>
                  <linearGradient id="investGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor="#fbbf24" stopOpacity={0.35} />
                    <stop offset="100%" stopColor="#fbbf24" stopOpacity={0} />
                  </linearGradient>
                </defs>
                <CartesianGrid stroke="#1e293b" vertical={false} />
                <XAxis dataKey="date" tick={{ fontSize: 10, fill: "#64748b" }} tickFormatter={(d) => new Date(d).toLocaleDateString("pl-PL", { month: "short", year: "2-digit" })} />
                <YAxis hide />
                <Tooltip
                  cursor={false}
                  formatter={(v) => fmtPLN(v)}
                  labelFormatter={(d) => new Date(d).toLocaleDateString("pl-PL")}
                  contentStyle={{ background: "#0f172a", border: "1px solid #1e293b", borderRadius: 8, fontSize: 12 }}
                      itemStyle={{ color: "#e2e8f0" }}
                      labelStyle={{ color: "#94a3b8" }}
                />
                <Area type="monotone" dataKey="value" stroke="#fbbf24" strokeWidth={2} fill="url(#investGrad)" isAnimationActive />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>
      ) : (
        <div className="rounded-2xl border border-dashed border-slate-800 p-6 text-center mb-6">
          <p className="text-sm text-slate-500">Za mało transakcji, żeby pokazać wykres w czasie.</p>
        </div>
      )}

      {categoryData.length > 0 ? (
        <div className="rounded-2xl bg-slate-900 border border-slate-800 p-4 mb-6">
          <p className="text-xs text-slate-400 mb-3">Podział portfela wg kategorii</p>
          <div style={{ height: 32 * categoryData.length + 20 }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={categoryData} layout="vertical" margin={{ left: 8, right: 40 }}>
                <XAxis type="number" hide domain={[0, (max) => max * 1.2]} />
                <YAxis type="category" dataKey="name" tick={{ fontSize: 11, fill: "#94a3b8" }} width={110} />
                <Tooltip
                  cursor={noHighlightCursor}
                  content={({ active, payload }) => {
                    if (!active || !payload || !payload.length) return null;
                    const d = payload[0].payload;
                    return (
                      <div style={{ background: "#0f172a", border: "1px solid #1e293b", borderRadius: 8, fontSize: 12, padding: "8px 10px", maxWidth: 220 }}>
                        <p style={{ color: "#e2e8f0", fontWeight: 600, marginBottom: 4 }}>{d.name} · {d.percent.toFixed(0)}%</p>
                        <p style={{ color: "#94a3b8" }}>{d.companies.join(", ")}</p>
                      </div>
                    );
                  }}
                />
                <Bar dataKey="value" radius={[0, 6, 6, 0]} isAnimationActive activeBar={false}>
                  {categoryData.map((d, i) => <Cell key={i} fill={CATEGORY_COLORS[d.name] || "#94a3b8"} />)}
                  <LabelList
                    dataKey="percent"
                    position="right"
                    formatter={(v) => `${v.toFixed(0)}%`}
                    style={{ fill: "#cbd5e1", fontSize: 11 }}
                  />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
          <p className="text-xs text-slate-600 mt-3">Dotknij słupka, żeby zobaczyć, jakie spółki się na niego składają.</p>
        </div>
      ) : (
        <div className="rounded-2xl border border-dashed border-slate-800 p-6 text-center mb-6">
          <p className="text-sm text-slate-500">Brak otwartych pozycji do pokazania podziału.</p>
        </div>
      )}

      <div className="rounded-2xl bg-slate-900 border border-slate-800 p-4 mb-6">
        <p className="text-xs text-slate-400 mb-3">Porównanie z benchmarkiem</p>
        <div className="flex flex-wrap gap-2 mb-3">
          {BENCHMARKS.map((b) => (
            <button
              key={b.key}
              onClick={() => toggleBenchmark(b.key)}
              className={`px-3 py-1.5 rounded-full text-xs font-medium transition-colors ${selectedBenchmarks.includes(b.key) ? "bg-amber-400 text-slate-900" : "bg-slate-800 text-slate-400 border border-slate-700"}`}
            >
              {b.label}
            </button>
          ))}
        </div>
        <button
          onClick={runComparison}
          disabled={!selectedBenchmarks.length || loadingBenchmarks || !filteredTransactions.length}
          className="w-full py-2.5 rounded-xl bg-slate-800 text-slate-200 text-sm font-medium disabled:opacity-40 mb-3"
        >
          {loadingBenchmarks ? "Porównuję…" : "Porównaj"}
        </button>

        {benchmarkError && (
          <p className="text-xs text-rose-400 mb-3">Nie udało się pobrać danych benchmarku: {benchmarkError}</p>
        )}

        {benchmarkSeries.length > 1 && (
          <div style={{ height: 220 }}>
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart data={benchmarkSeries} margin={{ left: 0, right: 8, top: 8, bottom: 0 }}>
                <CartesianGrid vertical={false} stroke="#1e293b" strokeDasharray="3 3" />
                <XAxis
                  dataKey="date"
                  tickFormatter={(d) => new Date(d).toLocaleDateString("pl-PL", { month: "short", year: "2-digit" })}
                  tick={{ fontSize: 10, fill: "#64748b" }}
                  axisLine={{ stroke: "#1e293b" }}
                  tickLine={false}
                  minTickGap={30}
                />
                <YAxis tick={{ fontSize: 10, fill: "#64748b" }} tickFormatter={(v) => fmtPLNShort(v)} width={48} />
                <Tooltip
                  cursor={{ stroke: "#475569" }}
                  labelFormatter={(d) => new Date(d).toLocaleDateString("pl-PL")}
                  formatter={(v) => fmtPLN(v)}
                  contentStyle={{ background: "#0f172a", border: "1px solid #1e293b", borderRadius: 8, fontSize: 12 }}
                  itemStyle={{ color: "#e2e8f0" }}
                  labelStyle={{ color: "#94a3b8", marginBottom: 4 }}
                />
                <Legend wrapperStyle={{ fontSize: 11 }} />
                <Line type="monotone" dataKey="Twój portfel" stroke={seriesColors[0]} strokeWidth={2.5} dot={false} isAnimationActive />
                {selectedBenchmarks.map((key, i) => {
                  const label = BENCHMARKS.find((b) => b.key === key).label;
                  return (
                    <Line
                      key={key}
                      type="monotone"
                      dataKey={label}
                      stroke={seriesColors[(i + 1) % seriesColors.length]}
                      strokeWidth={2}
                      dot={false}
                      isAnimationActive
                      connectNulls
                    />
                  );
                })}
              </ComposedChart>
            </ResponsiveContainer>
          </div>
        )}
        <p className="text-xs text-slate-600 mt-3">
          Wartość portfela (kapitał od transakcji, na koniec realna wycena) vs symulacja "gdyby te same wpłaty trafiły w benchmark". Wymaga zewnętrznych danych giełdowych (Yahoo Finance) — w środowisku Claude może się nie udać z powodu ograniczeń sieciowych.
        </p>
      </div>
    </div>
  );
}

const TIMEFRAMES = [
  { key: "1d", label: "1D", days: 1 },
  { key: "1t", label: "1T", days: 7 },
  { key: "1m", label: "1M", days: 30 },
  { key: "6m", label: "6M", days: 182 },
  { key: "1r", label: "1R", days: 365 },
  { key: "5l", label: "5L", days: 365 * 5 },
  { key: "max", label: "Max", days: null },
];

function StockDetailScreen({ ticker, transactions, portfolios, prices, twelveDataKey, onBack }) {
  const [historySeries, setHistorySeries] = useState(null);
  const [historyError, setHistoryError] = useState("");
  const [loadingHistory, setLoadingHistory] = useState(true);
  const [timeframe, setTimeframe] = useState("1r");

  const tickerTx = useMemo(
    () => transactions.filter((t) => t.ticker.toUpperCase() === ticker.toUpperCase()).sort((a, b) => new Date(a.date) - new Date(b.date)),
    [transactions, ticker]
  );

  const holdingInfo = useMemo(() => {
    let qty = 0, costBasis = 0, realized = 0;
    let name = ticker, portfolioIds = new Set();
    for (const t of tickerTx) {
      if (t.name?.trim()) name = t.name.trim();
      portfolioIds.add(t.portfolioId);
      const q = Number(t.quantity) || 0, p = Number(t.price) || 0;
      if (t.type === "buy") { qty += q; costBasis += q * p; }
      else {
        const avg = qty > 0 ? costBasis / qty : 0;
        const sq = Math.min(q, qty);
        realized += (p - avg) * sq;
        qty -= q; costBasis -= avg * sq;
        if (qty < 0) qty = 0;
        if (costBasis < 0) costBasis = 0;
      }
    }
    const avgCost = qty > 0 ? costBasis / qty : 0;
    return { name, qty, costBasis, avgCost, realized, portfolioIds: [...portfolioIds] };
  }, [tickerTx, ticker]);

  const priceInfo = prices[ticker.toUpperCase()];
  const currentPrice = priceInfo?.price ?? holdingInfo.avgCost;
  const currentValue = holdingInfo.qty * currentPrice;
  const unrealizedPnl = currentValue - holdingInfo.costBasis;
  const unrealizedPnlPercent = holdingInfo.costBasis > 0 ? (unrealizedPnl / holdingInfo.costBasis) * 100 : 0;
  const totalPnl = unrealizedPnl + holdingInfo.realized;
  const category = CATEGORY_MAP[ticker.toUpperCase()] || "Inne";
  const portfolioNames = holdingInfo.portfolioIds.map((id) => portfolios.find((p) => p.id === id)?.name).filter(Boolean);

  const cutoffDate = useMemo(() => {
    const tf = TIMEFRAMES.find((t) => t.key === timeframe);
    if (!tf || !tf.days) {
      return tickerTx.length ? new Date(tickerTx[0].date) : new Date(0);
    }
    return new Date(Date.now() - tf.days * 24 * 60 * 60 * 1000);
  }, [timeframe, tickerTx]);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoadingHistory(true);
      setHistoryError("");
      try {
        if (timeframe === "1d") {
          if (!twelveDataKey) throw new Error("Widok 1D wymaga klucza API Twelve Data (Ustawienia)");
          const isPL = ticker.toUpperCase().endsWith(".PL");
          const isUS = ticker.toUpperCase().endsWith(".US");
          const bareSymbol = ticker.replace(/\.(US|PL)$/i, "");
          const rows = await fetchTwelveDataIntraday(bareSymbol, isPL ? "XWAR" : null, twelveDataKey);
          if (!cancelled) setHistorySeries(rows.map((r) => ({ t: r.t, price: r.close })));
        } else {
          const tf = TIMEFRAMES.find((t) => t.key === timeframe);
          const rangeIntervalMap = {
            "1t": { range: "5d", interval: "30m" },
            "1m": { range: "1mo", interval: "1d" },
            "6m": { range: "6mo", interval: "1d" },
            "1r": { range: "1y", interval: "1d" },
            "5l": { range: "5y", interval: "1wk" },
            max: { range: "max", interval: "1mo" },
          };
          const cfg = rangeIntervalMap[timeframe] || { range: "3mo", interval: "1d" };
          const yahooSymbol = toYahooSymbol(ticker);
          const rows = await fetchYahooSeries(yahooSymbol, cfg.range, cfg.interval);
          if (!cancelled) setHistorySeries(rows.map((r) => ({ t: new Date(r.date).getTime(), price: r.close })));
        }
      } catch (e) {
        if (!cancelled) { setHistoryError(String(e.message || e)); setHistorySeries(null); }
      }
      if (!cancelled) setLoadingHistory(false);
    })();
    return () => { cancelled = true; };
  }, [ticker, timeframe, twelveDataKey]);

  const txPointsAll = tickerTx.map((t) => ({ t: new Date(t.date).getTime(), price: Number(t.price), type: t.type }));
  const cutoffMs = cutoffDate.getTime();
  const txPoints = txPointsAll.filter((p) => p.t >= cutoffMs);

  const historyInRange = historySeries ? historySeries.filter((p) => p.t >= cutoffMs) : null;
  const chartData = historyInRange && historyInRange.length > 1 ? historyInRange : txPoints.map((p) => ({ t: p.t, price: p.price }));
  const usingOwnData = !(historyInRange && historyInRange.length > 1);

  return (
    <div className="min-h-full w-full bg-slate-950 text-slate-100 font-sans">
      <style>{`.recharts-wrapper *:focus { outline: none !important; }`}</style>
      <div className="max-w-md mx-auto px-4 pt-6 pb-10">
        <div className="flex items-center gap-3 mb-5">
          <button onClick={onBack} className="w-9 h-9 rounded-full bg-slate-900 border border-slate-800 flex items-center justify-center shrink-0">
            <ArrowLeft size={18} />
          </button>
          <CompanyIcon ticker={ticker} size={36} />
          <div className="min-w-0">
            <p className="text-base font-bold truncate">{holdingInfo.name}</p>
            <p className="text-xs text-slate-500">{ticker}</p>
          </div>
        </div>

        <div className="rounded-2xl bg-slate-900 border border-slate-800 p-4 mb-4">
          <p className="text-2xl tabular-nums tracking-tight font-bold tracking-tight mb-1">{fmtPLN(currentPrice)}</p>
          {chartData.length > 1 && (() => {
            const periodStart = chartData[0].price;
            const periodEnd = chartData[chartData.length - 1].price;
            const periodChange = periodEnd - periodStart;
            const periodChangePct = periodStart ? (periodChange / periodStart) * 100 : 0;
            const trendUp = periodChange >= 0;
            const chartColor = trendUp ? "#34d399" : "#f87171";
            const tfLabel = TIMEFRAMES.find((t) => t.key === timeframe)?.label || "";
            return (
              <>
                <p className={`text-sm font-semibold mb-2 ${trendUp ? "text-emerald-400" : "text-rose-400"}`}>
                  {trendUp ? "+" : ""}{fmtPLN(periodChange)} · {trendUp ? "+" : ""}{periodChangePct.toFixed(2)}% · {tfLabel}
                </p>
                <div style={{ height: 180 }} className="mt-2">
                  <ResponsiveContainer width="100%" height="100%">
                    <ComposedChart data={chartData} margin={{ left: 0, right: 8, top: 8, bottom: 0 }}>
                      <defs>
                        <linearGradient id="priceGradient" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor={chartColor} stopOpacity={0.4} />
                          <stop offset="100%" stopColor={chartColor} stopOpacity={0} />
                        </linearGradient>
                      </defs>
                      <CartesianGrid vertical={false} stroke="#1e293b" strokeDasharray="3 3" />
                      <XAxis
                        dataKey="t"
                        type="number"
                        domain={["dataMin", "dataMax"]}
                        tickFormatter={(t) =>
                          timeframe === "1d" || timeframe === "1t"
                            ? new Date(t).toLocaleDateString("pl-PL", { day: "numeric", month: "short" })
                            : new Date(t).toLocaleDateString("pl-PL", { month: "short", day: "numeric" })
                        }
                        tick={{ fontSize: 10, fill: "#64748b" }}
                        axisLine={{ stroke: "#1e293b" }}
                        tickLine={false}
                      />
                      <YAxis hide domain={["auto", "auto"]} />
                      <Tooltip
                        cursor={{ stroke: "#475569", strokeWidth: 1 }}
                        content={({ active, payload, label }) => {
                          if (!active || !payload || !payload.length) return null;
                          const entry = payload.find((p) => p.name !== "marker") || payload[0];
                          return (
                            <div style={{ background: "#0f172a", border: "1px solid #1e293b", borderRadius: 8, fontSize: 12, padding: "6px 10px" }}>
                              <div style={{ color: "#94a3b8", marginBottom: 2 }}>
                                {new Date(label).toLocaleDateString("pl-PL", { day: "numeric", month: "short", year: "numeric" })}
                                {(timeframe === "1d" || timeframe === "1t") &&
                                  " · " + new Date(label).toLocaleTimeString("pl-PL", { hour: "2-digit", minute: "2-digit" })}
                              </div>
                              <div style={{ color: "#e2e8f0", fontWeight: 700 }}>{fmtPLN(entry.value)}</div>
                            </div>
                          );
                        }}
                      />
                      <Area
                        type="monotone"
                        dataKey="price"
                        stroke={chartColor}
                        strokeWidth={2.5}
                        fill="url(#priceGradient)"
                        dot={false}
                        isAnimationActive
                        activeDot={{ r: 4, fill: chartColor, stroke: "#0f172a", strokeWidth: 2 }}
                      />
                      <Scatter
                        name="marker"
                        data={txPoints}
                        dataKey="price"
                    shape={(p) => {
                      const isBuy = p.payload.type === "buy";
                      return (
                        <circle
                          cx={p.cx}
                          cy={p.cy}
                          r={5}
                          fill={isBuy ? "#34d399" : "#f87171"}
                          stroke="#0f172a"
                          strokeWidth={2}
                        />
                      );
                    }}
                  />
                </ComposedChart>
              </ResponsiveContainer>
            </div>
              </>
            );
          })()}
          <div className="flex items-center gap-1.5 mt-3">
            {TIMEFRAMES.map((tf) => (
              <button
                key={tf.key}
                onClick={() => setTimeframe(tf.key)}
                className={`px-3 py-1.5 rounded-full text-xs font-bold transition-colors ${timeframe === tf.key ? "bg-amber-400 text-slate-900" : "bg-slate-800 text-slate-400"}`}
              >
                {tf.label}
              </button>
            ))}
          </div>
          <div className="flex items-center gap-3 mt-3 text-xs text-slate-500">
            <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-emerald-400" /> Kupno</span>
            <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-rose-400" /> Sprzedaż</span>
            {loadingHistory && <span className="ml-auto">Ładuję historię…</span>}
          </div>
          {!loadingHistory && usingOwnData && (
            <p className="text-xs text-slate-600 mt-2">
              {timeframe === "1d" && historyError && historyError.includes("klucza API")
                ? historyError
                : "Pełna historia notowań chwilowo niedostępna — wykres pokazuje Twoje własne transakcje."}
            </p>
          )}
        </div>

        <div className="rounded-2xl bg-slate-900 border border-slate-800 p-4 mb-4">
          <p className="text-sm font-bold mb-3">Podsumowanie pozycji</p>
          <div className="grid grid-cols-2 gap-y-3 text-sm">
            <div>
              <p className="text-xs text-slate-500">Posiadana ilość</p>
              <p className="font-bold tabular-nums">{holdingInfo.qty}</p>
            </div>
            <div>
              <p className="text-xs text-slate-500">Śr. cena zakupu</p>
              <p className="font-bold tabular-nums">{fmtPLN(holdingInfo.avgCost)}</p>
            </div>
            <div>
              <p className="text-xs text-slate-500">Wartość obecna</p>
              <p className="font-bold tabular-nums">{fmtPLN(currentValue)}</p>
            </div>
            <div>
              <p className="text-xs text-slate-500">Zysk / strata</p>
              <p className={`font-bold tabular-nums ${totalPnl >= 0 ? "text-emerald-400" : "text-rose-400"}`}>
                {fmtPLN(totalPnl)} ({fmtPct(unrealizedPnlPercent)})
              </p>
            </div>
            <div>
              <p className="text-xs text-slate-500">Kategoria</p>
              <p className="font-bold">{category}</p>
            </div>
            <div>
              <p className="text-xs text-slate-500">Portfel</p>
              <p className="font-bold truncate">{portfolioNames.join(", ") || "—"}</p>
            </div>
          </div>
        </div>

        <div className="rounded-2xl bg-slate-900 border border-slate-800 p-4">
          <p className="text-sm font-bold mb-3">Historia transakcji</p>
          <div className="space-y-2">
            {[...tickerTx].reverse().map((t) => (
              <div key={t.id} className="flex items-center gap-3 text-sm">
                <span className={`w-2 h-2 rounded-full shrink-0 ${t.type === "buy" ? "bg-emerald-400" : "bg-rose-400"}`} />
                <span className="text-slate-400 flex-1">{t.type === "buy" ? "Kupno" : "Sprzedaż"} {t.quantity} szt.</span>
                <span className="font-bold tabular-nums">{fmtPLN(t.price)}</span>
                <span className="text-xs text-slate-500 w-16 text-right shrink-0">{new Date(t.date).toLocaleDateString("pl-PL")}</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

function Sheet({ onClose, children, wide }) {
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-black/60" onClick={onClose}>
      <div
        className={`w-full max-w-md bg-slate-950 border-t border-slate-800 rounded-t-3xl p-5 pb-8 ${wide ? "max-h-[85vh]" : "max-h-[75vh]"} overflow-y-auto animate-[slideUp_0.25s_ease-out]`}
        onClick={(e) => e.stopPropagation()}
      >
        <style>{`@keyframes slideUp{from{transform:translateY(100%)}to{transform:translateY(0)}}`}</style>
        {children}
      </div>
    </div>
  );
}

function TransactionForm({ portfolios, onManage, onClose, onSubmit }) {
  const [portfolioId, setPortfolioId] = useState(portfolios[0]?.id || "");
  const [ticker, setTicker] = useState("");
  const [name, setName] = useState("");
  const [type, setType] = useState("buy");
  const [quantity, setQuantity] = useState("");
  const [price, setPrice] = useState("");
  const [currency, setCurrency] = useState("PLN");
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [error, setError] = useState("");

  function handleSubmit() {
    if (!portfolioId) return setError("Dodaj najpierw portfel.");
    if (!ticker.trim()) return setError("Podaj ticker spółki.");
    if (!quantity || Number(quantity) <= 0) return setError("Podaj poprawną liczbę akcji.");
    if (!price || Number(price) <= 0) return setError("Podaj poprawną cenę.");
    if (!date) return setError("Podaj datę transakcji.");
    onSubmit({ portfolioId, ticker: ticker.trim().toUpperCase(), name: name.trim(), type, quantity: Number(quantity), price: Number(price), currency, date });
  }

  return (
    <Sheet onClose={onClose}>
      <div className="flex items-center justify-between mb-5">
        <h2 className="text-base font-bold">Nowa transakcja</h2>
        <button onClick={onClose} className="text-slate-500"><X size={20} /></button>
      </div>

      <div className="grid grid-cols-2 gap-2 mb-4">
        <button
          onClick={() => setType("buy")}
          className={`py-2.5 rounded-xl text-sm font-medium transition-colors ${type === "buy" ? "bg-emerald-500 text-slate-950" : "bg-slate-900 text-slate-400 border border-slate-800"}`}
        >
          Kupno
        </button>
        <button
          onClick={() => setType("sell")}
          className={`py-2.5 rounded-xl text-sm font-medium transition-colors ${type === "sell" ? "bg-rose-500 text-slate-950" : "bg-slate-900 text-slate-400 border border-slate-800"}`}
        >
          Sprzedaż
        </button>
      </div>

      <label className="block text-xs text-slate-400 mb-1.5 flex items-center gap-1"><Building2 size={12} /> Portfel</label>
      {portfolios.length === 0 ? (
        <button onClick={onManage} className="w-full mb-4 bg-slate-900 border border-dashed border-slate-700 rounded-xl px-3 py-2.5 text-sm text-amber-400">
          Dodaj pierwszy portfel
        </button>
      ) : (
        <div className="flex gap-2 mb-4">
          <select value={portfolioId} onChange={(e) => setPortfolioId(e.target.value)} className="flex-1 bg-slate-900 border border-slate-800 rounded-xl px-3 py-2.5 text-sm text-slate-100">
            {portfolios.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
          </select>
          <button onClick={onManage} className="w-11 shrink-0 bg-slate-900 border border-slate-800 rounded-xl flex items-center justify-center text-slate-400">
            <Settings size={15} />
          </button>
        </div>
      )}

      <div className="grid grid-cols-2 gap-3 mb-4">
        <div>
          <label className="block text-xs text-slate-400 mb-1.5">Ticker</label>
          <input value={ticker} onChange={(e) => setTicker(e.target.value)} placeholder="np. PKN" className="w-full bg-slate-900 border border-slate-800 rounded-xl px-3 py-2.5 text-sm text-slate-100 placeholder-slate-600" />
        </div>
        <div>
          <label className="block text-xs text-slate-400 mb-1.5">Nazwa (opcjonalnie)</label>
          <input value={name} onChange={(e) => setName(e.target.value)} placeholder="np. Orlen" className="w-full bg-slate-900 border border-slate-800 rounded-xl px-3 py-2.5 text-sm text-slate-100 placeholder-slate-600" />
        </div>
      </div>

      <div className="grid grid-cols-2 gap-3 mb-4">
        <div>
          <label className="block text-xs text-slate-400 mb-1.5 flex items-center gap-1"><Hash size={12} /> Liczba akcji</label>
          <input type="number" inputMode="decimal" value={quantity} onChange={(e) => setQuantity(e.target.value)} placeholder="0" className="w-full bg-slate-900 border border-slate-800 rounded-xl px-3 py-2.5 text-sm text-slate-100 font-mono placeholder-slate-600" />
        </div>
        <div>
          <label className="block text-xs text-slate-400 mb-1.5">Cena za akcję</label>
          <input type="number" inputMode="decimal" value={price} onChange={(e) => setPrice(e.target.value)} placeholder="0.00" className="w-full bg-slate-900 border border-slate-800 rounded-xl px-3 py-2.5 text-sm text-slate-100 font-mono placeholder-slate-600" />
        </div>
      </div>

      <div className="grid grid-cols-2 gap-3 mb-4">
        <div>
          <label className="block text-xs text-slate-400 mb-1.5">Waluta</label>
          <select value={currency} onChange={(e) => setCurrency(e.target.value)} className="w-full bg-slate-900 border border-slate-800 rounded-xl px-3 py-2.5 text-sm text-slate-100">
            {CURRENCIES.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <div>
          <label className="block text-xs text-slate-400 mb-1.5 flex items-center gap-1"><Calendar size={12} /> Data</label>
          <input type="date" value={date} onChange={(e) => setDate(e.target.value)} className="w-full bg-slate-900 border border-slate-800 rounded-xl px-3 py-2.5 text-sm text-slate-100" />
        </div>
      </div>

      {error && <p className="text-xs text-rose-400 mb-3">{error}</p>}

      <button onClick={handleSubmit} className="w-full bg-amber-400 text-slate-900 font-bold py-3 rounded-xl active:scale-[0.98] transition-transform">
        Dodaj transakcję
      </button>
    </Sheet>
  );
}

function HistorySheet({ transactions, portfolioName, onClose, onDelete }) {
  return (
    <Sheet onClose={onClose}>
      <div className="flex items-center justify-between mb-4">
        <h2 className="text-base font-bold">Historia transakcji</h2>
        <button onClick={onClose} className="text-slate-500"><X size={20} /></button>
      </div>
      <div className="space-y-2">
        {transactions.map((t) => (
          <div key={t.id} className="flex items-center gap-3 bg-slate-900 border border-slate-800 rounded-xl p-3">
            <div className={`w-2 h-2 rounded-full shrink-0 ${t.type === "buy" ? "bg-emerald-400" : "bg-rose-400"}`} />
            <div className="min-w-0 flex-1">
              <p className="text-sm font-medium">{t.ticker} <span className="text-slate-500 font-normal">· {portfolioName(t.portfolioId)}</span></p>
              <p className="text-xs text-slate-500 tabular-nums tracking-tight">
                {t.type === "buy" ? "Kupno" : "Sprzedaż"} {t.quantity} szt. × {fmtPLN(t.price)} · {new Date(t.date).toLocaleDateString("pl-PL")}
              </p>
            </div>
            <button onClick={() => onDelete(t.id)} className="text-slate-600 shrink-0">
              <Trash2 size={16} />
            </button>
          </div>
        ))}
      </div>
    </Sheet>
  );
}

const FIELD_DEFS = [
  { key: "ticker", label: "Ticker", required: true, hints: ["ticker", "symbol", "walor", "instrument", "papier"] },
  { key: "name", label: "Nazwa spółki", required: false, hints: ["nazwa", "opis", "papier"] },
  { key: "type", label: "Typ (kupno/sprzedaż)", required: false, hints: ["typ", "rodzaj", "operacja", "strona", "k/s"] },
  { key: "quantity", label: "Liczba akcji", required: true, hints: ["ilość", "ilosc", "liczba", "wolumen", "sztuk"] },
  { key: "price", label: "Cena", required: true, hints: ["cena", "kurs"] },
  { key: "currency", label: "Waluta", required: false, hints: ["waluta", "currency"] },
  { key: "date", label: "Data", required: true, hints: ["data"] },
];

function guessMapping(columns) {
  const mapping = {};
  const lower = columns.map((c) => c.toLowerCase());
  for (const f of FIELD_DEFS) {
    const idx = lower.findIndex((c) => f.hints.some((h) => c.includes(h)));
    mapping[f.key] = idx >= 0 ? columns[idx] : "";
  }
  return mapping;
}

function sheetToMatrix(ws) {
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
}

function findHeaderRow(matrix, requiredLabels) {
  for (let i = 0; i < matrix.length; i++) {
    const row = matrix[i].map((c) => String(c).trim().toLowerCase());
    if (requiredLabels.every((lbl) => row.some((c) => c === lbl.toLowerCase()))) return i;
  }
  return -1;
}

function rowsFromHeader(matrix, headerIdx) {
  const headers = matrix[headerIdx].map((c) => String(c).trim());
  const out = [];
  for (let i = headerIdx + 1; i < matrix.length; i++) {
    const raw = matrix[i];
    if (!raw || raw.every((c) => String(c).trim() === "")) continue;
    const obj = {};
    headers.forEach((h, j) => { if (h) obj[h] = raw[j]; });
    out.push(obj);
  }
  return out;
}

function isXtbWorkbook(wb) {
  const names = wb.SheetNames.map((n) => n.toLowerCase());
  return names.includes("closed positions") || names.includes("open positions");
}

function parsePLNumber(raw) {
  const cleaned = String(raw).replace(/[\s\u00a0]/g, "").replace(",", ".");
  const n = parseFloat(cleaned);
  return Number.isFinite(n) ? n : NaN;
}

function detectIngMaklerskiCsv(text) {
  const firstLines = text.split(/\r?\n/).filter((l) => l.trim()).slice(0, 5);
  if (!firstLines.length) return false;
  let matches = 0;
  for (const line of firstLines) {
    const cols = line.split(";");
    if (cols.length === 9 && /^\d{2}-\d{2}-\d{4} \d{2}:\d{2}:\d{2}$/.test(cols[0].trim())) {
      matches++;
    }
  }
  return matches === firstLines.length;
}

function parseIngMaklerskiCsv(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.trim());
  const txs = [];
  for (const line of lines) {
    const cols = line.split(";");
    if (cols.length !== 9) continue;
    const [dateRaw, , tickerRaw, typeRaw, qtyRaw, priceRaw] = cols;
    const [d, m, y] = dateRaw.split(" ")[0].split("-");
    const time = dateRaw.split(" ")[1] || "00:00:00";
    const date = `${y}-${m}-${d}T${time}`;
    const ticker = tickerRaw.trim().toUpperCase() + ".PL";
    const type = typeRaw.trim().toLowerCase() === "kupno" ? "buy" : "sell";
    const quantity = parsePLNumber(qtyRaw);
    const price = parsePLNumber(priceRaw);
    if (!ticker || !Number.isFinite(quantity) || !Number.isFinite(price)) continue;
    txs.push({ ticker, name: ticker.replace(/\.PL$/, ""), type, quantity, price, currency: "PLN", date });
  }
  return txs;
}

function parseXtbWorkbook(wb) {
  const txs = [];
  const warnings = [];

  const closedSheet = wb.SheetNames.find((n) => n.toLowerCase() === "closed positions");
  if (closedSheet) {
    const matrix = sheetToMatrix(wb.Sheets[closedSheet]);
    const hIdx = findHeaderRow(matrix, ["Instrument", "Ticker", "Volume"]);
    if (hIdx >= 0) {
      for (const r of rowsFromHeader(matrix, hIdx)) {
        const category = String(r["Category"] || "").trim().toUpperCase();
        if (category !== "STOCK" && category !== "ETF") continue;
        const ticker = String(r["Ticker"] || "").trim().toUpperCase();
        const volume = parseNum(r["Volume"]);
        const purchaseValue = parseNum(r["Purchase Value"]);
        const saleValue = parseNum(r["Sale Value"]);
        if (!ticker || !Number.isFinite(volume) || volume <= 0) continue;
        const openDate = normalizeDate(r["Open Time (UTC)"]);
        const closeDate = normalizeDate(r["Close Time (UTC)"]);
        const name = String(r["Instrument"] || ticker).trim();
        if (Number.isFinite(purchaseValue) && purchaseValue > 0 && openDate) {
          txs.push({ ticker, name, type: "buy", quantity: volume, price: purchaseValue / volume, currency: "PLN", date: openDate });
        }
        if (Number.isFinite(saleValue) && saleValue > 0 && closeDate) {
          txs.push({ ticker, name, type: "sell", quantity: volume, price: saleValue / volume, currency: "PLN", date: closeDate });
        }
      }
    }
  }

  const openSheet = wb.SheetNames.find((n) => n.toLowerCase() === "open positions");
  if (openSheet) {
    const matrix = sheetToMatrix(wb.Sheets[openSheet]);
    const hIdx = findHeaderRow(matrix, ["Ticker", "Volume", "Open price"]);
    if (hIdx >= 0) {
      const openRows = rowsFromHeader(matrix, hIdx);
      const nameByTicker = {};
      for (const r of openRows) {
        const tk = String(r["Ticker"] || "").trim().toUpperCase();
        const label = String(r["Instrument/Position"] || "").trim();
        if (tk && label && !/^\d+$/.test(label)) nameByTicker[tk] = label;
      }
      for (const r of openRows) {
        const ticker = String(r["Ticker"] || "").trim().toUpperCase();
        const type = String(r["Type"] || "").trim().toUpperCase();
        if (!ticker || type !== "BUY") continue;
        const volume = parseNum(r["Volume"]);
        const value = parseNum(r["Value"]);
        const netProfit = parseNum(r["Net Profit"]);
        const currentPrice = parseNum(r["Current price"]);
        const openPrice = parseNum(r["Open price"]);
        const openDate = normalizeDate(r["Open time (UTC)"]);
        if (!Number.isFinite(volume) || volume <= 0 || !openDate) continue;
        let pricePLN = NaN;
        if (Number.isFinite(value) && Number.isFinite(netProfit)) {
          pricePLN = (value - netProfit) / volume;
        } else if (Number.isFinite(value) && Number.isFinite(currentPrice) && currentPrice > 0 && Number.isFinite(openPrice)) {
          const fx = value / volume / currentPrice;
          pricePLN = openPrice * fx;
        } else if (Number.isFinite(value)) {
          pricePLN = value / volume;
        }
        if (!Number.isFinite(pricePLN) || pricePLN <= 0) continue;
        const label = String(r["Instrument/Position"] || "").trim();
        txs.push({
          ticker,
          name: nameByTicker[ticker] || (/^\d+$/.test(label) ? ticker : label),
          type: "buy",
          quantity: volume,
          price: pricePLN,
          currency: "PLN",
          date: openDate,
        });
      }
    }
  }

  if (!txs.length) warnings.push("Nie znaleziono transakcji w arkuszach XTB.");
  return { txs, warnings };
}

function ImportModal({ portfolios, onAddPortfolio, onClose, onImport }) {
  const [fileName, setFileName] = useState("");
  const [columns, setColumns] = useState([]);
  const [rows, setRows] = useState([]);
  const [mapping, setMapping] = useState({});
  const [autoTxs, setAutoTxs] = useState(null);
  const [portfolioChoice, setPortfolioChoice] = useState(portfolios[0]?.id || "__new__");
  const [newPortfolioName, setNewPortfolioName] = useState("");
  const [error, setError] = useState("");
  const [importSummary, setImportSummary] = useState(null);

  function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    setImportSummary(null);
    setAutoTxs(null);
    setFileName(file.name);
    const ext = file.name.split(".").pop().toLowerCase();
    const reader = new FileReader();
    if (ext === "csv") {
      reader.onload = (ev) => {
        const buf = ev.target.result;
        let text = new TextDecoder("utf-8", { fatal: false }).decode(buf);
        if (text.includes("\ufffd")) {
          text = new TextDecoder("windows-1250", { fatal: false }).decode(buf);
        }
        if (detectIngMaklerskiCsv(text)) {
          const txs = parseIngMaklerskiCsv(text);
          if (txs.length) {
            setAutoTxs(txs);
            setColumns(["__auto__"]);
            return;
          }
        }
        const result = Papa.parse(text, { header: true, skipEmptyLines: true });
        const cols = result.meta.fields || [];
        setColumns(cols);
        setRows(result.data);
        setMapping(guessMapping(cols));
      };
      reader.readAsArrayBuffer(file);
    } else if (ext === "xlsx" || ext === "xls") {
      reader.onload = (ev) => {
        try {
          const wb = XLSX.read(ev.target.result, { type: "array", cellDates: false, raw: true });

          if (isXtbWorkbook(wb)) {
            const { txs, warnings } = parseXtbWorkbook(wb);
            if (txs.length) {
              setAutoTxs(txs);
              setColumns(["__auto__"]);
              return;
            }
            setError(warnings[0] || "Nie znaleziono transakcji w pliku XTB.");
            return;
          }

          const sheet = wb.Sheets[wb.SheetNames[0]];
          const matrix = sheetToMatrix(sheet);
          let headerIdx = 0;
          let best = 0;
          for (let i = 0; i < Math.min(matrix.length, 15); i++) {
            const filled = matrix[i].filter((c) => String(c).trim() !== "").length;
            if (filled > best) { best = filled; headerIdx = i; }
          }
          const parsed = rowsFromHeader(matrix, headerIdx);
          const cols = matrix[headerIdx].map((c) => String(c).trim()).filter(Boolean);
          if (!cols.length) { setError("Nie udało się rozpoznać kolumn w pliku."); return; }
          setColumns(cols);
          setRows(parsed);
          setMapping(guessMapping(cols));
        } catch (err) {
          setError("Nie udało się odczytać pliku Excel.");
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      setError("Obsługiwane formaty: CSV, XLSX, XLS.");
    }
  }

  function buildRows() {
    let portfolioId = portfolioChoice;
    if (portfolioChoice === "__new__") {
      if (!newPortfolioName.trim()) { setError("Podaj nazwę nowego portfela."); return null; }
      const p = onAddPortfolio(newPortfolioName.trim());
      portfolioId = p.id;
    }
    const built = [];
    let skipped = 0;
    for (const row of rows) {
      const ticker = mapping.ticker ? String(row[mapping.ticker] || "").trim().toUpperCase() : "";
      const quantity = mapping.quantity ? parseNum(row[mapping.quantity]) : NaN;
      const price = mapping.price ? parseNum(row[mapping.price]) : NaN;
      const date = mapping.date ? normalizeDate(row[mapping.date]) : "";
      if (!ticker || !Number.isFinite(quantity) || quantity <= 0 || !Number.isFinite(price) || price <= 0 || !date) {
        skipped++;
        continue;
      }
      const rawType = mapping.type ? String(row[mapping.type]).toLowerCase() : "";
      const type = /sprzed|sell/.test(rawType) ? "sell" : "buy";
      const currency = mapping.currency ? (String(row[mapping.currency]).trim().toUpperCase() || "PLN") : "PLN";
      const name = mapping.name ? String(row[mapping.name] || "").trim() : "";
      built.push({ portfolioId, ticker, name, type, quantity, price, currency, date });
    }
    return { built, skipped };
  }

  function resolvePortfolioId() {
    if (portfolioChoice === "__new__") {
      if (!newPortfolioName.trim()) { setError("Podaj nazwę nowego portfela."); return null; }
      return onAddPortfolio(newPortfolioName.trim()).id;
    }
    return portfolioChoice;
  }

  function handleImport() {
    setError("");
    if (autoTxs) {
      const portfolioId = resolvePortfolioId();
      if (!portfolioId) return;
      const duplicates = onImport(autoTxs.map((t) => ({ ...t, portfolioId })));
      setImportSummary({ imported: autoTxs.length - duplicates, duplicates, invalid: 0 });
      return;
    }
    const result = buildRows();
    if (!result) return;
    if (result.built.length === 0) {
      setError("Nie znaleziono poprawnych wierszy do zaimportowania. Sprawdź mapowanie kolumn.");
      return;
    }
    const duplicates = onImport(result.built);
    setImportSummary({ imported: result.built.length - duplicates, duplicates, invalid: result.skipped });
  }

  return (
    <Sheet onClose={onClose} wide>
      <div className="flex items-center justify-between mb-4">
        <h2 className="text-base font-bold">Import z pliku</h2>
        <button onClick={onClose} className="text-slate-500"><X size={20} /></button>
      </div>

      {importSummary ? (
        <div className="text-center py-4">
          <div className="w-14 h-14 rounded-full bg-emerald-950 flex items-center justify-center mx-auto mb-4">
            <Check size={26} className="text-emerald-400" />
          </div>
          <p className="text-base font-bold mb-1">Zaimportowano {importSummary.imported} transakcji</p>
          {(importSummary.duplicates > 0 || importSummary.invalid > 0) && (
            <p className="text-sm text-slate-400 mb-1">
              {importSummary.duplicates > 0 && `Pominięto ${importSummary.duplicates} duplikatów (już były w portfelu)`}
              {importSummary.duplicates > 0 && importSummary.invalid > 0 && " · "}
              {importSummary.invalid > 0 && `${importSummary.invalid} niepoprawnych wierszy`}
            </p>
          )}
          <button onClick={onClose} className="mt-4 w-full bg-amber-400 text-slate-900 font-bold py-3 rounded-xl active:scale-[0.98] transition-transform">
            Gotowe
          </button>
        </div>
      ) : !columns.length ? (
        <>
          <p className="text-sm text-slate-300 mb-1">Wybierz plik CSV lub XLSX</p>
          <p className="text-xs text-slate-500 mb-4">wyciąg transakcji z ING, XTB, IKE, PZU TFI</p>
          <input
            type="file"
            accept=".csv,.xlsx,.xls"
            onChange={handleFile}
            className="block w-full text-sm text-slate-300 mb-4 file:mr-3 file:py-3 file:px-5 file:rounded-xl file:border-0 file:text-sm file:font-bold file:bg-amber-400 file:text-slate-900"
          />
          {error && <p className="text-xs text-rose-400 mt-3">{error}</p>}
        </>
      ) : autoTxs ? (
        <>
          <div className="rounded-xl bg-emerald-950/40 border border-emerald-900/60 p-3 mb-4">
            <p className="text-sm text-emerald-300 mb-1">Rozpoznano format XTB</p>
            <p className="text-xs text-emerald-500">{fileName} · {autoTxs.length} transakcji, ceny przeliczone na PLN</p>
          </div>

          <div className="mb-4 rounded-xl border border-slate-800 overflow-hidden">
            <div className="bg-slate-900 px-3 py-2 text-xs text-slate-500">Podgląd</div>
            <div className="divide-y divide-slate-800 max-h-48 overflow-y-auto">
              {autoTxs.map((t, i) => (
                <div key={i} className="px-3 py-2 flex items-center gap-2 text-xs">
                  <span className={`w-1.5 h-1.5 rounded-full shrink-0 ${t.type === "buy" ? "bg-emerald-400" : "bg-rose-400"}`} />
                  <span className="font-medium text-slate-200">{t.ticker}</span>
                  <span className="text-slate-500 tabular-nums tracking-tight truncate">{t.quantity} × {fmtPLN(t.price)}</span>
                  <span className="text-slate-600 ml-auto shrink-0">{t.date}</span>
                </div>
              ))}
            </div>
          </div>

          <p className="text-xs text-slate-400 mb-2">Przypisz do portfela</p>
          <select value={portfolioChoice} onChange={(e) => setPortfolioChoice(e.target.value)} className="w-full mb-4 bg-slate-900 border border-slate-800 rounded-xl px-3 py-2.5 text-sm text-slate-100">
            {portfolios.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
            <option value="__new__">+ Nowy portfel</option>
          </select>
          {portfolioChoice === "__new__" && (
            <input
              value={newPortfolioName}
              onChange={(e) => setNewPortfolioName(e.target.value)}
              placeholder="Nazwa nowego portfela"
              className="w-full mb-4 bg-slate-900 border border-slate-800 rounded-xl px-3 py-2.5 text-sm text-slate-100 placeholder-slate-600"
            />
          )}

          {error && <p className="text-xs text-rose-400 mb-3">{error}</p>}

          <button onClick={handleImport} className="w-full bg-amber-400 text-slate-900 font-bold py-3 rounded-xl active:scale-[0.98] transition-transform">
            Importuj {autoTxs.length} transakcji
          </button>
        </>
      ) : (
        <>
          <p className="text-xs text-slate-500 mb-3">{fileName} · {rows.length} wierszy</p>

          <p className="text-xs text-slate-400 mb-2">Dopasuj kolumny</p>
          <div className="grid grid-cols-2 gap-2 mb-4">
            {FIELD_DEFS.map((f) => (
              <div key={f.key}>
                <label className="block text-xs text-slate-500 mb-1">{f.label}{f.required ? " *" : ""}</label>
                <select
                  value={mapping[f.key] || ""}
                  onChange={(e) => setMapping((m) => ({ ...m, [f.key]: e.target.value }))}
                  className="w-full bg-slate-900 border border-slate-800 rounded-lg px-2 py-2 text-xs text-slate-100"
                >
                  <option value="">— brak —</option>
                  {columns.map((c) => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
            ))}
          </div>

          <p className="text-xs text-slate-400 mb-2">Przypisz do portfela</p>
          <div className="flex gap-2 mb-4">
            <select value={portfolioChoice} onChange={(e) => setPortfolioChoice(e.target.value)} className="flex-1 bg-slate-900 border border-slate-800 rounded-xl px-3 py-2.5 text-sm text-slate-100">
              {portfolios.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
              <option value="__new__">+ Nowy portfel</option>
            </select>
          </div>
          {portfolioChoice === "__new__" && (
            <input
              value={newPortfolioName}
              onChange={(e) => setNewPortfolioName(e.target.value)}
              placeholder="Nazwa nowego portfela"
              className="w-full mb-4 bg-slate-900 border border-slate-800 rounded-xl px-3 py-2.5 text-sm text-slate-100 placeholder-slate-600"
            />
          )}

          {rows.length > 0 && (
            <div className="mb-4 rounded-xl border border-slate-800 overflow-hidden">
              <div className="bg-slate-900 px-3 py-2 text-xs text-slate-500">Podgląd pierwszych wierszy</div>
              <div className="divide-y divide-slate-800">
                {rows.slice(0, 3).map((r, i) => (
                  <div key={i} className="px-3 py-2 text-xs tabular-nums tracking-tight text-slate-400 truncate">
                    {mapping.ticker ? r[mapping.ticker] : "?"} · {mapping.quantity ? r[mapping.quantity] : "?"} szt. × {mapping.price ? r[mapping.price] : "?"}
                  </div>
                ))}
              </div>
            </div>
          )}

          {error && <p className="text-xs text-rose-400 mb-3">{error}</p>}

          <button onClick={handleImport} className="w-full bg-amber-400 text-slate-900 font-bold py-3 rounded-xl active:scale-[0.98] transition-transform">
            Importuj transakcje
          </button>
        </>
      )}
    </Sheet>
  );
}
