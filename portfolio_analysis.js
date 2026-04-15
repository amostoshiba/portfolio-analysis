const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  TableOfContents
} = require('docx');
const fs = require('fs');

// ─── Colour palette ───────────────────────────────────────────────────────────
const C = {
  NAVY:       "1E3A5F",
  BLUE:       "2E75B6",
  LIGHT_BLUE: "D5E8F0",
  MID_BLUE:   "BDD5E8",
  GREEN:      "1A7A3F",
  GREEN_BG:   "C8EDD6",
  AMBER:      "8B5A00",
  AMBER_BG:   "FFF0C0",
  RED:        "A00000",
  RED_BG:     "FADADD",
  GRAY_BG:    "F2F2F2",
  DISASTER_BG:"F5C0C0",
  EXPECTED_BG:"D5E8F0",
  WHITE:      "FFFFFF",
  BORDER:     "AAAAAA",
};

const border = { style: BorderStyle.SINGLE, size: 1, color: C.BORDER };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorders = {
  top: { style: BorderStyle.NONE, size: 0, color: C.WHITE },
  bottom: { style: BorderStyle.NONE, size: 0, color: C.WHITE },
  left: { style: BorderStyle.NONE, size: 0, color: C.WHITE },
  right: { style: BorderStyle.NONE, size: 0, color: C.WHITE },
};
const cellPad = { top: 80, bottom: 80, left: 120, right: 120 };

// ─── Helpers ──────────────────────────────────────────────────────────────────
const sp = (n) => new Paragraph({ children: [new TextRun("")], spacing: { after: n } });

const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun({ text, font: "Arial", size: 32, bold: true, color: C.NAVY })],
  spacing: { before: 320, after: 160 },
});

const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children: [new TextRun({ text, font: "Arial", size: 26, bold: true, color: C.BLUE })],
  spacing: { before: 240, after: 120 },
});

const h3 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  children: [new TextRun({ text, font: "Arial", size: 22, bold: true, color: C.NAVY })],
  spacing: { before: 160, after: 80 },
});

const para = (text, opts = {}) => new Paragraph({
  children: [new TextRun({ text, font: "Arial", size: 20, ...opts })],
  spacing: { after: 120 },
});

const paraRuns = (runs, spacingAfter = 120) => new Paragraph({
  children: runs.map(r => new TextRun({ font: "Arial", size: 20, ...r })),
  spacing: { after: spacingAfter },
});

const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

const cell = (text, fill, textColor = "000000", bold = false, width = 1560, italic = false) =>
  new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: cellPad, verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      children: [new TextRun({ text, font: "Arial", size: 18, bold, color: textColor, italic })],
      alignment: AlignmentType.LEFT,
    })],
  });

const cellC = (text, fill, textColor = "000000", bold = false, width = 1560) =>
  new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: cellPad, verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      children: [new TextRun({ text, font: "Arial", size: 18, bold, color: textColor })],
      alignment: AlignmentType.CENTER,
    })],
  });

const verdictBadge = (verdict) => {
  const map = { GREEN: [C.GREEN, C.GREEN_BG], AMBER: [C.AMBER, C.AMBER_BG], RED: [C.RED, C.RED_BG] };
  const [tc, bg] = map[verdict] || ["000000", C.WHITE];
  return new Paragraph({
    children: [new TextRun({ text: `  ▶  ${verdict}  `, font: "Arial", size: 22, bold: true, color: tc, shading: { type: ShadingType.CLEAR, fill: bg } })],
    spacing: { after: 120 },
  });
};

// ─── Probability-weighted scenario table ──────────────────────────────────────
// 4 rows: Disaster, Bear, Base, Bull + 1 Expected row
const dcfTable = (rows, expectedCagr, probOfLoss) => {
  const W = [1200, 640, 1360, 1480, 1200, 1060, 1280];  // 8220 total
  const hdrs = ["Scenario", "Probability", "Annual Growth", "Terminal Multiple", "10yr CAGR", "10yr Total Return", "15% CAGR Entry"];
  const hdrRow = new TableRow({
    children: hdrs.map((h, i) => new TableCell({
      borders, width: { size: W[i], type: WidthType.DXA },
      shading: { fill: C.NAVY, type: ShadingType.CLEAR },
      margins: cellPad,
      children: [new Paragraph({ children: [new TextRun({ text: h, font: "Arial", size: 16, bold: true, color: C.WHITE })], alignment: AlignmentType.CENTER })],
    })),
  });

  const scenarioFills = {
    "Disaster": C.DISASTER_BG,
    "Bear": C.AMBER_BG,
    "Base": C.LIGHT_BLUE,
    "Bull": C.GREEN_BG,
  };

  const totalReturn = (cagr) => {
    const r = Math.pow(1 + cagr / 100, 10) - 1;
    const pct = Math.round(r * 100);
    return pct >= 0 ? `+${pct}%` : `${pct}%`;
  };

  const dataRows = rows.map((r) => {
    const fill = scenarioFills[r.scenario] || C.GRAY_BG;
    const cagrColor = r.cagr >= 15 ? C.GREEN : r.cagr >= 10 ? C.AMBER : C.RED;
    const trColor = r.cagr >= 0 ? C.GREEN : C.RED;
    const scenarioLabel = r.scenario === "Disaster" ? "☠  Disaster" : r.scenario;
    // Highlight entry price: green if current price is at or below, amber if within 20% above
    const entryText = r.buyPrice || "—";
    return new TableRow({
      children: [
        cell(scenarioLabel, fill, r.scenario === "Disaster" ? C.RED : "000000", true, W[0]),
        cellC(r.prob, fill, "444444", false, W[1]),
        cellC(r.growth, fill, "000000", false, W[2]),
        cellC(r.multiple, fill, "000000", false, W[3]),
        cellC(`${r.cagr.toFixed(1)}%`, fill, cagrColor, true, W[4]),
        cellC(totalReturn(r.cagr), fill, trColor, true, W[5]),
        cellC(entryText, fill, r.buyPrice ? C.NAVY : "999999", r.buyPrice ? true : false, W[6]),
      ],
    });
  });

  // Expected weighted row
  const expColor = expectedCagr >= 15 ? C.GREEN : expectedCagr >= 10 ? C.AMBER : C.RED;
  const expTrColor = expectedCagr >= 0 ? C.GREEN : C.RED;
  const expectedRow = new TableRow({
    children: [
      cell("Expected (weighted)", C.MID_BLUE, C.NAVY, true, W[0]),
      cellC("100%", C.MID_BLUE, C.NAVY, false, W[1]),
      cellC("—", C.MID_BLUE, "555555", false, W[2]),
      cellC("—", C.MID_BLUE, "555555", false, W[3]),
      cellC(`${expectedCagr.toFixed(1)}%`, C.MID_BLUE, expColor, true, W[4]),
      cellC(totalReturn(expectedCagr), C.MID_BLUE, expTrColor, true, W[5]),
      cellC("—", C.MID_BLUE, "555555", false, W[6]),
    ],
  });

  return new Table({
    width: { size: 8220, type: WidthType.DXA },
    columnWidths: W,
    rows: [hdrRow, ...dataRows, expectedRow],
  });
};

// ─── Metrics row table ────────────────────────────────────────────────────────
const metricsTable = (items) => {
  const W_L = 1200, W_V = 900;
  const cols = 4;
  const makeRow = (slice) => new TableRow({
    children: slice.map(item => [
      new TableCell({
        borders, width: { size: W_L, type: WidthType.DXA },
        shading: { fill: C.NAVY, type: ShadingType.CLEAR }, margins: cellPad,
        children: [new Paragraph({ children: [new TextRun({ text: item.label, font: "Arial", size: 16, bold: true, color: C.WHITE })] })],
      }),
      new TableCell({
        borders, width: { size: W_V, type: WidthType.DXA },
        shading: { fill: C.LIGHT_BLUE, type: ShadingType.CLEAR }, margins: cellPad,
        children: [new Paragraph({ children: [new TextRun({ text: item.value, font: "Arial", size: 16, bold: false, color: "000000" })] })],
      }),
    ]).flat(),
  });
  const row1 = items.slice(0, cols);
  const row2 = items.slice(cols, cols * 2);
  while (row1.length < cols) row1.push({ label: "", value: "" });
  while (row2.length < cols) row2.push({ label: "", value: "" });
  const totalW = (W_L + W_V) * cols;
  return new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: Array(cols).fill(null).flatMap(() => [W_L, W_V]),
    rows: [makeRow(row1), makeRow(row2)],
  });
};

// ─── Risk badge ───────────────────────────────────────────────────────────────
const riskPara = (risk, text) => {
  const map = { "VERY LOW": [C.GREEN, C.GREEN_BG], "LOW": [C.GREEN, C.GREEN_BG], "LOW–MEDIUM": [C.AMBER, C.AMBER_BG], "MEDIUM": [C.AMBER, C.AMBER_BG], "MEDIUM–HIGH": [C.RED, C.RED_BG], "HIGH": [C.RED, C.RED_BG] };
  const [tc, bg] = map[risk] || [C.AMBER, C.AMBER_BG];
  return new Paragraph({
    children: [
      new TextRun({ text: `Disruption Risk: `, font: "Arial", size: 20, bold: true }),
      new TextRun({ text: ` ${risk} `, font: "Arial", size: 20, bold: true, color: tc, shading: { type: ShadingType.CLEAR, fill: bg } }),
      new TextRun({ text: `  —  ${text}`, font: "Arial", size: 20, italic: true }),
    ],
    spacing: { after: 120 },
  });
};

// ─── Divider ──────────────────────────────────────────────────────────────────
const divider = () => new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.BLUE, space: 1 } },
  spacing: { after: 160 },
  children: [],
});

// ─────────────────────────────────────────────────────────────────────────────
//  COMPANY DATA  (4 scenarios each: Disaster / Bear / Base / Bull)
// ─────────────────────────────────────────────────────────────────────────────

const companies = [

  {
    name: "Games Workshop (GAW)",
    ticker: "LON: GAW",
    currentPrice: "~£170",
    verdict: "AMBER",
    expectedCagr: 8.8,
    probOfLoss: 12,
    metrics: [
      { label: "Price", value: "~£170" },
      { label: "TTM EPS", value: "~£5.75" },
      { label: "Trailing PE", value: "~29x" },
      { label: "ROCE", value: "191%" },
      { label: "Net Margin", value: "35%" },
      { label: "Debt", value: "Zero" },
      { label: "Rev Growth (H1 FY26)", value: "+34% YoY" },
      { label: "EPS Growth (FY25)", value: "+30%" },
    ],
    quality: [
      "Games Workshop is the creator and licensor of the Warhammer IP universe — a 40-year-old franchise with extraordinary brand loyalty and a self-reinforcing community ecosystem. With 191% ROCE, zero debt, and a 35% net margin, few businesses of any size produce more cash per unit of invested capital. The moat is multi-layered: deep collector psychology, network effects within hobby communities, and a lore archive expensive and time-consuming to replicate.",
      "The Warhammer IP is entering a major licensing phase: Amazon TV projects in development, dozens of video game releases per year, and board game adaptations. This licensing revenue carries near-zero marginal cost and falls almost entirely to the bottom line. Management has behaved like genuine owner-operators — resisting cheap licensing, investing in IP quality, and returning cash through dividends and buybacks.",
    ],
    risk: "LOW–MEDIUM",
    riskText: "Franchise fatigue if core IP loses cultural relevance; key-man risk on creative leadership; digital gaming reducing appetite for physical miniatures among younger audiences; China expansion failing to materialise; Amazon TV show underperforming.",
    dcf: [
      { scenario: "Disaster", prob: "12%", growth: "EPS flat (0%/yr)", multiple: "10x PE (franchise collapse)", cagr: -10.3 },
      { scenario: "Bear",     prob: "25%", growth: "10% EPS/yr",       multiple: "20x PE",                    cagr: 5.8,  buyPrice: "£74"  },
      { scenario: "Base",     prob: "45%", growth: "15% EPS/yr",       multiple: "22x PE",                    cagr: 11.7, buyPrice: "£127" },
      { scenario: "Bull",     prob: "18%", growth: "20% EPS/yr",       multiple: "25x PE",                    cagr: 18.0, buyPrice: "£220" },
    ],
    verdict_text: "At ~£170, GAW's probability-weighted expected CAGR is ~8.8% — decent but not compelling at this price. The 12% disaster probability (franchise fatigue, digital substitution of younger audiences) is the honest assessment: it has happened to comparable hobby IP franchises before. The business is genuinely exceptional and the bull case (18% CAGR) is achievable given recent 34% revenue growth, but this is not the right price to add aggressively. Your existing position at ~£100 cost basis is excellent — let it compound. The sweet spot for adding more is £127–£145.",
    addZone: "£127–£145",
  },

  {
    name: "MercadoLibre (MELI)",
    ticker: "NASDAQ: MELI",
    currentPrice: "~$1,660",
    verdict: "GREEN",
    expectedCagr: 14.1,
    probOfLoss: 12,
    metrics: [
      { label: "Price", value: "~$1,660" },
      { label: "Reported EPS", value: "~$41" },
      { label: "Owner Earnings/sh", value: "~$59 (est.)" },
      { label: "P/Owner Earnings", value: "~28x" },
      { label: "Rev Growth (FY24)", value: "+38%" },
      { label: "EPS Growth TTM", value: "+44%" },
      { label: "Net Margin", value: "~10% (by design)" },
      { label: "Debt", value: "Moderate (fintech)" },
    ],
    quality: [
      "MercadoLibre is the dominant e-commerce and fintech platform across Latin America — 600M+ people with structurally under-penetrated digital commerce. MELI runs at low reported margins by deliberate choice, reinvesting aggressively into logistics, credit, and ecosystem lock-in — mirroring early Amazon. Owner earnings of ~$3bn reflect a business generating far more true cash than the income statement shows. As reinvestment matures, margins will expand sharply, producing an earnings step-change.",
      "The competitive moat has three reinforcing layers: network effects between buyers and sellers, the MercadoPago ecosystem creating deep financial switching costs, and logistics infrastructure that took $3bn+ to build. The business is arguably the single highest-quality compounder in the portfolio on a growth-adjusted basis.",
    ],
    risk: "MEDIUM",
    riskText: "LatAm political risk is the key tail risk — Brazil capital controls, regulatory crackdown, or currency devaluation could impair returns dramatically. This is genuine and warrants careful position sizing. Competitive risk from Amazon and regional players is real but manageable.",
    dcf: [
      { scenario: "Disaster", prob: "12%", growth: "OE –5%/yr",         multiple: "10x P/OE (LatAm crisis)",   cagr: -14.8 },
      { scenario: "Bear",     prob: "23%", growth: "15% OE/yr",          multiple: "20x P/OE",                  cagr: 11.1,  buyPrice: "$1,180" },
      { scenario: "Base",     prob: "45%", growth: "20% OE/yr",          multiple: "25x P/OE",                  cagr: 18.6,  buyPrice: "$2,257" },
      { scenario: "Bull",     prob: "20%", growth: "25% OE/yr",          multiple: "28x P/OE",                  cagr: 24.9,  buyPrice: "$3,802" },
    ],
    verdict_text: "At ~$1,660, MELI has a probability-weighted expected CAGR of ~14.1% — the second-highest in the portfolio. The 12% disaster probability reflects genuine LatAm political risk that cannot be dismissed; this is why position sizing discipline matters. But the base case (20% owner earnings growth, well below current trajectory) gives 18.6% CAGR, and even the bear case (15% growth) gives 11.1%. Your existing position at ~$1,920 is at a higher cost but the thesis is intact and the business is more attractively priced today. GREEN — consider a small add if it pulls back toward $1,400–$1,500.",
    addZone: "$1,400–$1,700",
  },

  {
    name: "Constellation Software / Topicus",
    ticker: "TSX: CSU / TSXV: TOI",
    currentPrice: "CSU ~$1,800 USD (~$2,500 CAD) / TOI ~$98 CAD",
    verdict: "GREEN",
    expectedCagr: 17.2,
    probOfLoss: 5,
    metrics: [
      { label: "CSU Price", value: "~$2,500 CAD (~$1,800 USD)" },
      { label: "TOI Price", value: "~$98 CAD" },
      { label: "CSU Rev CAGR (5yr)", value: "+26%" },
      { label: "ROIC", value: ">20%" },
      { label: "CSU FCF/sh (est.)", value: "~$81 USD" },
      { label: "P/FCF (est.)", value: "~22x" },
      { label: "TOI Rev Growth FY25", value: "+20%" },
      { label: "Debt", value: "Moderate (acquisitions)" },
    ],
    quality: [
      "Constellation Software is arguably one of the greatest capital allocation businesses ever created. Since Mark Leonard founded it in 1995, CSU has acquired and operated hundreds of vertical market software (VMS) businesses — mission-critical, niche software serving regulated industries (utilities, healthcare, logistics, municipalities). Churn rates are under 2% annually. The reported PE (50–95x) is meaningless — it reflects amortisation of acquired intangibles. The correct metric is FCF: ~$1.7bn USD, giving a P/FCF of ~22x (USD market cap $38bn / USD FCF $1.7bn).",
      "Topicus is CSU's European spinout applying the same model to European VMS. Mark Leonard's annual letters are among the most honest, intelligent shareholder communications in existence. The disaster scenario is the only real risk: the acquisition machine breaking down through overpayment or cultural degradation post-Leonard. The 5% probability reflects this being real but unlikely given the systems-based culture built over 30 years.",
    ],
    risk: "VERY LOW",
    riskText: "VMS businesses have near-permanent customer lock-in. The main risk is CSU running out of acquisition targets or overpaying systematically. Both are monitored explicitly by management. Scale is a genuine headwind but Topicus, international expansion, and smaller sub-acquisitions are the clear management response.",
    dcf: [
      { scenario: "Disaster", prob: "5%",  growth: "FCF –2%/yr",  multiple: "12x P/FCF (model breakdown)",  cagr: -8.0  },
      { scenario: "Bear",     prob: "25%", growth: "12% FCF/yr",  multiple: "22x P/FCF",                    cagr: 11.9,  buyPrice: "$1,368 USD" },
      { scenario: "Base",     prob: "50%", growth: "18% FCF/yr",  multiple: "25x P/FCF",                    cagr: 19.4,  buyPrice: "$2,619 USD" },
      { scenario: "Bull",     prob: "20%", growth: "22% FCF/yr",  multiple: "28x P/FCF",                    cagr: 24.8,  buyPrice: "$4,095 USD" },
    ],
    verdict_text: "The highest expected CAGR in the portfolio at ~17.2%, with only a 5% probability of loss. At ~$1,800 USD (~22x P/FCF), the current price is well inside the base-case max buy of $2,619 USD, implying ~19.4% CAGR on the base scenario. Even the bear case (12% FCF/yr) gives 11.9%. This is the portfolio's single most compelling risk/reward position: world-class business, reasonable valuation, low loss probability, high expected return. Strong GREEN. Treat CSU and TOI as a combined position with CSU as the anchor.",
    addZone: "$1,600–$1,900 USD / $2,200–$2,600 CAD (CSU) / $90–$100 CAD (TOI)",
  },

  {
    name: "S&P Global (SPGI)",
    ticker: "NYSE: SPGI",
    currentPrice: "~$426",
    verdict: "AMBER",
    expectedCagr: 9.7,
    probOfLoss: 5,
    metrics: [
      { label: "Price", value: "~$426" },
      { label: "2025 EPS (GAAP)", value: "$14.66" },
      { label: "2026 Guided EPS", value: "$19.40–19.65" },
      { label: "Forward PE (2026)", value: "~22x" },
      { label: "Trailing PE", value: "~29x" },
      { label: "Rev Growth FY25", value: "+8%" },
      { label: "EPS Growth FY25", value: "+19%" },
      { label: "Debt", value: "Low (investment grade)" },
    ],
    quality: [
      "S&P Global is one of the most defensible businesses on earth. In credit ratings, it shares a global duopoly with Moody's — protected not by patents but by the regulatory structure of global capital markets. Every bond issuance globally requires a rating from S&P or Moody's. This is structural and has persisted for over 100 years. Beyond ratings: S&P Dow Jones Indices (S&P 500 licensor), Platts (commodity benchmarks), and Market Intelligence. The indices business is extraordinary — every new passive dollar flowing into markets generates recurring licensing revenue.",
      "The disaster scenario at 5% probability reflects a regulatory break-up separating the ratings and data businesses, combined with a severe multiple compression. This has been discussed in policy circles but faces enormous practical barriers given how embedded SPGI is in global capital market infrastructure.",
    ],
    risk: "VERY LOW",
    riskText: "The most serious risk — disintermediation of credit ratings by AI or private credit growth — remains distant and structural. Bond markets depend on standardised ratings. No credible substitute exists within a 10-year horizon.",
    dcf: [
      { scenario: "Disaster", prob: "5%",  growth: "EPS +5%/yr",  multiple: "12x PE (regulatory break-up)", cagr: -3.9  },
      { scenario: "Bear",     prob: "30%", growth: "9% EPS/yr",   multiple: "18x PE",                       cagr: 6.9,   buyPrice: "$205" },
      { scenario: "Base",     prob: "50%", growth: "12% EPS/yr",  multiple: "20x PE",                       cagr: 11.0,  buyPrice: "$299" },
      { scenario: "Bull",     prob: "15%", growth: "15% EPS/yr",  multiple: "23x PE",                       cagr: 15.6,  buyPrice: "$449" },
    ],
    verdict_text: "SPGI at ~$426 delivers a probability-weighted expected CAGR of ~9.7% — solid but not exceptional. The business is outstanding; the issue is entirely valuation. It re-rated from ~$337 to ~$426 over recent months and the ideal buying window (22x forward PE was more attractive at lower prices) has narrowed. The disaster probability is very low at 5%. AMBER — start a small initial tranche, but reserve the bulk of capital for weakness toward $380–$400 where the base case CAGR improves meaningfully.",
    addZone: "$380–$410 (current $426 is fair but stretched)",
  },

  {
    name: "Ferrari (RACE)",
    ticker: "NYSE: RACE",
    currentPrice: "~$319",
    verdict: "AMBER",
    expectedCagr: 8.1,
    probOfLoss: 10,
    metrics: [
      { label: "Price", value: "~$319" },
      { label: "FY25 EPS (EUR)", value: "€8.96" },
      { label: "FY25 EPS (USD est.)", value: "~$9.50" },
      { label: "Trailing PE", value: "~33x" },
      { label: "52-wk High", value: "$519" },
      { label: "Revenue Growth FY25", value: "+7%" },
      { label: "EPS Growth FY25", value: "+6%" },
      { label: "EBIT Margin FY25", value: "29.5%" },
    ],
    quality: [
      "Ferrari is not an automotive company. It is a luxury goods and IP licensing business that happens to produce cars. Its pricing power is virtually unlimited — it has a waitlist of wealthy buyers, and artificially constraining supply is core to the brand strategy. At 29.5% EBIT margin, it is one of the highest-quality industrial businesses in the world. The moat is cultural and historic: 75 years of racing heritage, an association with aspirational success that cannot be manufactured, and a collector market that treats older Ferraris as appreciating assets.",
      "The electrification transition is a genuine but manageable challenge — Ferrari's Formula 1 team already runs hybrid power units. The personalisation business (bespoke options) generates growing high-margin revenue. The disaster scenario: EV transition failure causing the brand to lose relevance with under-40s, with margins reverting toward automotive industry norms.",
    ],
    risk: "LOW–MEDIUM",
    riskText: "Electrification risk to the brand/performance experience is real but being managed. The deeper risk is valuation — at 33x PE with only 6-7% EPS growth, the numbers are difficult for a 15% return target without heroic assumptions.",
    dcf: [
      { scenario: "Disaster", prob: "10%", growth: "EPS flat (0%/yr)", multiple: "15x PE (EV/brand failure)",    cagr: -7.7  },
      { scenario: "Bear",     prob: "25%", growth: "8% EPS/yr",        multiple: "25x PE",                       cagr: 4.9,   buyPrice: "$127" },
      { scenario: "Base",     prob: "50%", growth: "12% EPS/yr",       multiple: "30x PE",                       cagr: 10.7,  buyPrice: "$219" },
      { scenario: "Bull",     prob: "15%", growth: "15% EPS/yr",       multiple: "35x PE",                       cagr: 15.5,  buyPrice: "$333" },
    ],
    verdict_text: "At ~$319, Ferrari has a probability-weighted expected CAGR of ~8.1% — below the portfolio average. The business is genuinely exceptional but the valuation is the problem: 33x PE with only 6% current EPS growth. The stock has fallen 38% from its $519 high, which makes it feel tempting — but the base-case buy price for 15% CAGR is ~$219, and even the bull case max buy is only $333. AMBER — the math is getting closer. Watch for a pullback to $240–$260 before building a full position. A small initial tranche at $300–$320 is defensible for a long-term holder who believes in the bull case.",
    addZone: "$220–$260 (full position) / $300–$320 (small initial tranche only)",
  },

  {
    name: "LVMH (MC)",
    ticker: "EPA: MC",
    currentPrice: "~€474",
    verdict: "RED",
    expectedCagr: 3.7,
    probOfLoss: 37,
    metrics: [
      { label: "Price", value: "~€474" },
      { label: "EPS (est. FY25)", value: "~€21" },
      { label: "Trailing PE", value: "~22x" },
      { label: "Revenue FY24", value: "€84.7bn" },
      { label: "Revenue FY25", value: "€80.8bn (↓)" },
      { label: "Op. Margin FY24", value: "23.1%" },
      { label: "Revenue Growth FY25", value: "–4.6% (organic –1%)" },
      { label: "Debt", value: "Manageable" },
    ],
    quality: [
      "LVMH is the world's premier luxury conglomerate — 75 houses including Louis Vuitton, Dior, Moët & Chandon, Hennessy, and Bulgari. The portfolio is genuinely irreplicable and Bernard Arnault has shown exceptional capital allocation discipline over 35 years. The core moat is desire and aspiration — LVMH's best brands operate where price increases enhance rather than reduce demand.",
      "However, the bear case is the honest case right now: FY2025 revenue declined from €84.7bn to €80.8bn. The Chinese luxury consumption slowdown is structural, not merely cyclical. Generation Z consumers in China are demonstrably less attracted to Western luxury brands than their parents. Succession risk around Arnault (77) is real. The business is excellent but the note of caution is strong.",
    ],
    risk: "MEDIUM",
    riskText: "China luxury structural slowdown; younger Chinese consumers de-prioritising Western luxury; aspirational luxury losing ground to 'experience economy'; FX headwinds from EUR strength; Arnault succession uncertainty.",
    dcf: [
      { scenario: "Disaster", prob: "12%", growth: "EPS –3%/yr",         multiple: "10x PE (China + succession)", cagr: -10.6 },
      { scenario: "Bear",     prob: "25%", growth: "EPS +2%/yr",          multiple: "16x PE (stagnation)",         cagr: -1.5  },
      { scenario: "Base",     prob: "48%", growth: "8% EPS/yr (recovery)",multiple: "20x PE",                      cagr: 6.7,   buyPrice: "€224" },
      { scenario: "Bull",     prob: "15%", growth: "13% EPS/yr",          multiple: "25x PE",                      cagr: 14.2,  buyPrice: "€440" },
    ],
    verdict_text: "LVMH is the weakest position in the portfolio on a probability-weighted basis. Expected CAGR of just ~3.7% — barely above inflation. More importantly, the probability of a negative 10-year return is ~37%: the disaster (12%) and bear (25%) scenarios both produce negative CAGRs. Revenue is actively declining. Even the bull case (13% growth, 25x PE) gives only 14.2% CAGR. This is a world-class business in a cyclical and structural downturn at a price that still embeds too much optimism. RED — do not add at current levels. The mathematically honest entry for even the base case to work is €280–€360.",
    addZone: "€280–€360 (do not buy above €400)",
  },

  {
    name: "Copart (CPRT)",
    ticker: "NASDAQ: CPRT",
    currentPrice: "~$33",
    verdict: "AMBER",
    expectedCagr: 8.2,
    probOfLoss: 15,
    metrics: [
      { label: "Price", value: "~$33" },
      { label: "FY25 EPS", value: "$1.61" },
      { label: "Trailing PE", value: "~22x" },
      { label: "EV/EBIT", value: "~16x" },
      { label: "ROIC", value: "32.5%" },
      { label: "Revenue Growth FY25", value: "+9.7%" },
      { label: "EPS Growth FY25", value: "+13.4%" },
      { label: "Debt", value: "Very low" },
    ],
    quality: [
      "Copart is the world's largest online vehicle auction platform, processing salvage and total-loss vehicles. The business is a 40-year network of physical auction yards (200+ globally) that took decades to build — irreplicable due to zoning, permitting, and land costs. The moat compounds through network effects: more buyers attract more sellers, who achieve higher prices, attracting more buyers. Over 750,000 registered buyers across 170+ countries.",
      "The disaster scenario deserves honest confrontation: autonomous vehicles reducing accident rates by 60%+ within the next decade is technologically credible. A 15% probability is assigned — not because it is imminent, but because it is directionally real. The business has 15–20 years before this becomes acute, and EVs are paradoxically increasing total losses today (high battery costs make EV repair economics negative). The disaster probability is the primary reason this is AMBER and not GREEN.",
    ],
    risk: "LOW–MEDIUM",
    riskText: "Autonomous vehicles reducing accident volumes is the primary long-term tail risk (15% disaster probability). Near-term EV dynamics are net positive — EVs create more total losses than ICE vehicles. Management has decades of runway to pivot into commercial vehicles and other auction categories.",
    dcf: [
      { scenario: "Disaster", prob: "15%", growth: "EPS –5%/yr",  multiple: "10x PE (AV disruption)",  cagr: -11.6 },
      { scenario: "Bear",     prob: "30%", growth: "9% EPS/yr",   multiple: "18x PE",                  cagr: 7.7,   buyPrice: "$17"  },
      { scenario: "Base",     prob: "40%", growth: "13% EPS/yr",  multiple: "21x PE",                  cagr: 13.2,  buyPrice: "$28"  },
      { scenario: "Bull",     prob: "15%", growth: "15% EPS/yr",  multiple: "23x PE",                  cagr: 16.4,  buyPrice: "$37"  },
    ],
    verdict_text: "At ~$33, Copart has a probability-weighted expected CAGR of ~8.2%. The 15% disaster probability (AV disruption) is the honest assessment and is the main drag on the expected return. Setting the disaster aside, the base/bull scenarios are genuinely attractive — 13.2% and 16.4% respectively. This is still an excellent business at a reasonable price. AMBER with above-average tail risk. Size the position to reflect this: a maximum of 7–8% of portfolio rather than the 9–10% you might allocate to CSU or Amazon.",
    addZone: "$26–$37 (current ~$33 is in range — start building)",
  },

  {
    name: "Mastercard (MA)",
    ticker: "NYSE: MA",
    currentPrice: "~$507",
    verdict: "AMBER",
    expectedCagr: 10.7,
    probOfLoss: 8,
    metrics: [
      { label: "Price", value: "~$507" },
      { label: "TTM EPS", value: "$16.54" },
      { label: "Trailing PE", value: "~30x" },
      { label: "ROIC", value: "50%" },
      { label: "Revenue Growth FY25", value: "+16.4%" },
      { label: "EPS Growth (forecast)", value: "~15–17%/yr" },
      { label: "Net Margin", value: "~45%" },
      { label: "Debt", value: "Low" },
    ],
    quality: [
      "Mastercard is one of the highest-quality businesses in the world — a capital-light payments network with 50% ROIC, 45% net margins, and structural tailwinds from the multi-decade shift from cash to digital payments. The network itself is the moat: 4.6bn cards, 100M merchant locations, 210 countries. Neither side of the network switches because the other side is already there. Unlike banks, Mastercard bears no credit risk — pure-play on global consumer spending growth with dramatically lower risk.",
      "The disaster scenario at 8% probability reflects the credible risk of real-time account-to-account payments (FedNow, PIX, UPI-style systems) bypassing card networks in major markets over the next decade. This is happening gradually in several markets already. The probability is higher than it was five years ago, which is worth acknowledging.",
    ],
    risk: "LOW–MEDIUM",
    riskText: "The most credible disruption is real-time A2A payments bypassing card rails in major markets. This is a 10–15 year structural risk management is responding to through acquisitions and partnerships. Crypto and stablecoin adoption is a secondary wild card.",
    dcf: [
      { scenario: "Disaster", prob: "8%",  growth: "EPS flat (0%/yr)", multiple: "12x PE (payments bypass)", cagr: -9.0  },
      { scenario: "Bear",     prob: "22%", growth: "10% EPS/yr",       multiple: "22x PE",                   cagr: 6.4,   buyPrice: "$233" },
      { scenario: "Base",     prob: "50%", growth: "15% EPS/yr",       multiple: "26x PE",                   cagr: 13.0,  buyPrice: "$430" },
      { scenario: "Bull",     prob: "20%", growth: "18% EPS/yr",       multiple: "30x PE",                   cagr: 17.7,  buyPrice: "$642" },
    ],
    verdict_text: "At ~$507, Mastercard has a probability-weighted expected CAGR of ~10.7%. The business is world-class with only an 8% loss probability, but at 30x PE it is modestly above the base-case 15% CAGR entry point (~$430). The expected return is solid rather than compelling. AMBER — this is a business you want to own for life, but current prices slightly reduce the potential upside. Start a smaller initial tranche at current levels; reserve the larger add for pullbacks toward $415–$430.",
    addZone: "$410–$430 (ideal) / $490–$510 (small tranche only)",
  },

  {
    name: "Amazon (AMZN)",
    ticker: "NASDAQ: AMZN",
    currentPrice: "~$210",
    verdict: "GREEN",
    expectedCagr: 15.9,
    probOfLoss: 5,
    metrics: [
      { label: "Price", value: "~$210" },
      { label: "TTM EPS", value: "~$7.29" },
      { label: "Trailing PE", value: "~29x" },
      { label: "Op. Income FY25", value: "$80bn (+16.6%)" },
      { label: "EPS Growth FY25", value: "+28.8%" },
      { label: "AWS Revenue Growth", value: "~20%+ ongoing" },
      { label: "Advertising Growth", value: "~18–20%" },
      { label: "Debt", value: "Manageable" },
    ],
    quality: [
      "Amazon is difficult to value using PE because it operates at low reported margins while building multi-generational infrastructure. The correct lens is operating income: $80bn in FY2025, growing 16.6%. Three distinct engines each with long runways: AWS (cloud/AI infrastructure), Advertising ($50bn+ run rate, targets buyers at point-of-purchase intent), and e-commerce (scale moat with logistics flywheel).",
      "AWS is arguably the most important piece of digital infrastructure built in the 21st century. Switching costs of migrating enterprise workloads are enormous. AWS's AI buildout (Bedrock, Trainium, Inferentia) positions it to capture a disproportionate share of AI infrastructure spend. The disaster scenario at 5% reflects antitrust break-up or severe cloud commoditisation — possible but unlikely given AWS's decade-long head start.",
    ],
    risk: "LOW",
    riskText: "Amazon is more disruptor than disrupted. Primary risks are regulatory (antitrust in e-commerce/AWS) and competitive (Google Cloud, Microsoft Azure). Neither is existential over a 10-year horizon. AWS AI positioning reduces competitive risk meaningfully.",
    dcf: [
      { scenario: "Disaster", prob: "5%",  growth: "EPS flat (0%/yr)", multiple: "12x PE (antitrust/cloud)", cagr: -8.4  },
      { scenario: "Bear",     prob: "25%", growth: "15% EPS/yr",       multiple: "20x PE",                  cagr: 10.9,  buyPrice: "$146" },
      { scenario: "Base",     prob: "50%", growth: "20% EPS/yr",       multiple: "23x PE",                  cagr: 17.4,  buyPrice: "$257" },
      { scenario: "Bull",     prob: "20%", growth: "25% EPS/yr",       multiple: "28x PE",                  cagr: 24.5,  buyPrice: "$470" },
    ],
    verdict_text: "At ~$210, Amazon has the second-highest expected CAGR in the portfolio at ~15.9%, with only a 5% loss probability. The current price is below the base-case max buy of $257 (20% EPS growth, 23x terminal PE), meaning the base case already implies 17.4% CAGR from here. Recent EPS growth of 28.8% in FY25 makes the 20% base case conservative. This is the clearest and largest immediate buy in the portfolio. Strong GREEN — deploy Tranche 1 now, add aggressively on any weakness.",
    addZone: "$146–$260 (currently at $210 — strong buy zone)",
  },

  {
    name: "Microsoft (MSFT)",
    ticker: "NASDAQ: MSFT",
    currentPrice: "~$399",
    verdict: "AMBER",
    expectedCagr: 13.5,
    probOfLoss: 3,
    metrics: [
      { label: "Price", value: "~$399" },
      { label: "TTM EPS", value: "~$16.05" },
      { label: "Trailing PE", value: "~25x" },
      { label: "Revenue Growth TTM", value: "+16.7%" },
      { label: "EPS Growth TTM", value: "+28.8%" },
      { label: "Azure Growth", value: "~30%+ ongoing" },
      { label: "AI Annual Run Rate", value: "$13bn (+175% YoY)" },
      { label: "Debt", value: "Low / net cash" },
    ],
    quality: [
      "Microsoft is a compound machine with three interlocking growth engines: Azure, Microsoft 365 + Copilot, and Gaming/LinkedIn. Each benefits from the others — Azure hosts Microsoft 365, which generates data that improves Copilot, which drives Azure consumption. Enterprise switching costs from the Microsoft ecosystem are enormous. The $13bn AI run rate growing 175% YoY is early-stage monetisation with a very long runway.",
      "The disaster scenario at only 3% probability is the lowest in the portfolio — reflecting Microsoft's position as an AI enabler rather than AI disruption victim. The risk is open-source commoditisation of software development reducing enterprise software spending, but Microsoft's distribution and ecosystem depth make this a slow-burn risk at best.",
    ],
    risk: "LOW",
    riskText: "Microsoft is positioned to benefit from most technological disruption scenarios. Antitrust risk is real but manageable given the broad enterprise dependency. The main risk is AI commoditising the margins of its software businesses — a slow-burn concern rather than a near-term disruption.",
    dcf: [
      { scenario: "Disaster", prob: "3%",  growth: "EPS +5%/yr",  multiple: "12x PE (open source/AI commodity)", cagr: -2.4  },
      { scenario: "Bear",     prob: "25%", growth: "12% EPS/yr",  multiple: "20x PE",                            cagr: 9.6,   buyPrice: "$246" },
      { scenario: "Base",     prob: "50%", growth: "15% EPS/yr",  multiple: "23x PE",                            cagr: 14.1,  buyPrice: "$369" },
      { scenario: "Bull",     prob: "22%", growth: "18% EPS/yr",  multiple: "26x PE",                            cagr: 18.5,  buyPrice: "$540" },
    ],
    verdict_text: "At ~$399, Microsoft has an expected CAGR of ~13.5% — the lowest loss probability (3%) in the entire portfolio. The base case (15% EPS growth, 23x terminal) gives 14.1% CAGR just below the 15% target, and with current EPS growth at 28.8% the bull case is very reachable. The main objection is not quality or valuation — it is that Microsoft, at ~$3 trillion market cap, has the most limited upside ceiling in the portfolio simply due to size. AMBER on valuation grounds — modestly above the ideal entry of $365–$375. Still worth a Tranche 1 at current levels.",
    addZone: "$354–$375 (ideal) / current $399 (acceptable Tranche 1)",
  },

];

// ─────────────────────────────────────────────────────────────────────────────
//  PORTFOLIO SUMMARY TABLE (probability-weighted)
// ─────────────────────────────────────────────────────────────────────────────
const summaryTable = () => {
  const W = [1700, 1100, 1000, 1500, 1260, 1400, 800]; // 8760 total
  const hdrs = ["Company", "Price", "Multiple", "Exp. CAGR (wtd)", "Prob. of Loss", "Add Zone", "Signal"];
  const hdrRow = new TableRow({
    children: hdrs.map((h, i) => new TableCell({
      borders, width: { size: W[i], type: WidthType.DXA },
      shading: { fill: C.NAVY, type: ShadingType.CLEAR }, margins: cellPad,
      children: [new Paragraph({ children: [new TextRun({ text: h, font: "Arial", size: 16, bold: true, color: C.WHITE })], alignment: AlignmentType.CENTER })],
    })),
  });

  const rows_data = [
    { name: "Games Workshop",    price: "~£170",                    pe: "29x PE",     exp: "8.8%",  loss: "12%",  zone: "£127–£145",              v: "AMBER" },
    { name: "MercadoLibre",      price: "~$1,660",                  pe: "28x P/OE",   exp: "14.1%", loss: "12%",  zone: "$1,400–$1,700",          v: "GREEN" },
    { name: "CSU / Topicus",     price: "$1,800 USD (~$2,500 CAD)", pe: "22x P/FCF",  exp: "17.2%", loss: "5%",   zone: "$1,600–$1,900 USD",      v: "GREEN" },
    { name: "S&P Global",        price: "~$426",                    pe: "22x fwd PE", exp: "9.7%",  loss: "5%",   zone: "$380–$410",              v: "AMBER" },
    { name: "Ferrari",           price: "~$319",                    pe: "33x PE",     exp: "8.1%",  loss: "10%",  zone: "$220–$260",              v: "AMBER" },
    { name: "LVMH",              price: "~€474",                    pe: "22x PE",     exp: "3.7%",  loss: "37%",  zone: "€280–€360",              v: "RED"   },
    { name: "Copart",            price: "~$33",                     pe: "22x PE",     exp: "8.2%",  loss: "15%",  zone: "$26–$37",                v: "AMBER" },
    { name: "Mastercard",        price: "~$507",                    pe: "30x PE",     exp: "10.7%", loss: "8%",   zone: "$410–$430",              v: "AMBER" },
    { name: "Amazon",            price: "~$210",                    pe: "29x PE",     exp: "15.9%", loss: "5%",   zone: "$146–$260",              v: "GREEN" },
    { name: "Microsoft",         price: "~$399",                    pe: "25x PE",     exp: "13.5%", loss: "3%",   zone: "$354–$375",              v: "AMBER" },
  ];

  const verdictColor = { GREEN: [C.GREEN, C.GREEN_BG], AMBER: [C.AMBER, C.AMBER_BG], RED: [C.RED, C.RED_BG] };
  const expColor = (s) => {
    const n = parseFloat(s);
    return n >= 14 ? C.GREEN : n >= 8 ? C.AMBER : C.RED;
  };
  const lossColor = (s) => {
    const n = parseFloat(s);
    return n >= 25 ? C.RED : n >= 12 ? C.AMBER : C.GREEN;
  };

  const dataRows = rows_data.map((r, idx) => {
    const fill = idx % 2 === 0 ? C.GRAY_BG : C.WHITE;
    const [vc, vbg] = verdictColor[r.v];
    return new TableRow({
      children: [
        cell(r.name, fill, "000000", true, W[0]),
        cellC(r.price, fill, "000000", false, W[1]),
        cellC(r.pe, fill, "000000", false, W[2]),
        cellC(r.exp, fill, expColor(r.exp), true, W[3]),
        cellC(r.loss, fill, lossColor(r.loss), true, W[4]),
        cell(r.zone, fill, "000000", false, W[5]),
        cellC(r.v, vbg, vc, true, W[6]),
      ],
    });
  });

  // Portfolio expected row
  const portfolioRow = new TableRow({
    children: [
      cell("PORTFOLIO AVG", C.NAVY, C.WHITE, true, W[0]),
      cellC("—", C.NAVY, C.WHITE, false, W[1]),
      cellC("—", C.NAVY, C.WHITE, false, W[2]),
      cellC("11.0%", C.NAVY, "7FFFAA", true, W[3]),
      cellC("~8.5% avg", C.NAVY, "FFD580", true, W[4]),
      cellC("—", C.NAVY, C.WHITE, false, W[5]),
      cellC("—", C.NAVY, C.WHITE, false, W[6]),
    ],
  });

  return new Table({ width: { size: 8760, type: WidthType.DXA }, columnWidths: W, rows: [hdrRow, ...dataRows, portfolioRow] });
};

// ─────────────────────────────────────────────────────────────────────────────
//  BUILD DOCUMENT
// ─────────────────────────────────────────────────────────────────────────────

const children = [];

// ── TITLE PAGE ────────────────────────────────────────────────────────────────
children.push(
  sp(2000),
  new Paragraph({
    children: [new TextRun({ text: "Concentrated Portfolio Analysis", font: "Arial", size: 52, bold: true, color: C.NAVY })],
    alignment: AlignmentType.CENTER, spacing: { after: 200 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "10-Stock Quality Portfolio  ·  March 2026", font: "Arial", size: 28, color: C.BLUE })],
    alignment: AlignmentType.CENTER, spacing: { after: 200 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "Probability-Weighted Return Analysis  ·  4-Scenario Framework  ·  Honest Expected CAGR", font: "Arial", size: 22, color: "555555", italic: true })],
    alignment: AlignmentType.CENTER, spacing: { after: 400 },
  }),
  divider(),
  new Paragraph({
    children: [new TextRun({ text: "Analytical framework only. Not financial advice. All figures derived from publicly available data.", font: "Arial", size: 18, italic: true, color: "888888" })],
    alignment: AlignmentType.CENTER, spacing: { after: 2000 },
  }),
  pageBreak(),
);

// ── EXECUTIVE SUMMARY ─────────────────────────────────────────────────────────
children.push(
  h1("Executive Summary"),
  divider(),
  para("This document takes a probability-weighted approach to each of the ten portfolio companies. Rather than a simple bear/base/bull framework — which tends to produce unrealistically positive results — each company is assessed across four scenarios including a Disaster case (business impairment, not just slow growth). Probabilities are assigned to each scenario and a weighted expected CAGR is calculated. This gives an honest picture of what the portfolio is likely to return.", { color: "333333" }),
  para("The key finding: the portfolio's probability-weighted expected CAGR is approximately 11% per year — well above a 5% floor but below the aspirational 15% target for most positions. Three companies are GREEN (CSU, Amazon, MELI with expected CAGR above 14%). Six are AMBER. One is RED (LVMH), which has a 37% probability of delivering a negative 10-year return at current prices. The probability of the whole portfolio failing to deliver 5%+ over 10 years is estimated at roughly 8–10%.", { color: "333333" }),
  sp(80),
  summaryTable(),
  sp(160),
  para("METHODOLOGY NOTE: Each company has four scenarios — Disaster (business impairment), Bear (slow growth), Base (moderate growth), Bull (strong growth). Probabilities sum to 100%. The weighted expected CAGR uses these probabilities. 'Prob of Loss' is the summed probability of scenarios producing a negative 10-year return. MELI and CSU use owner earnings / FCF per share rather than reported EPS. All prices are approximate as of March 2026.", { italic: true, color: "666666" }),
  pageBreak(),
);

// ── COMPANY SECTIONS ──────────────────────────────────────────────────────────
companies.forEach((co, idx) => {
  children.push(
    new Paragraph({
      children: [
        new TextRun({ text: co.name, font: "Arial", size: 36, bold: true, color: C.NAVY }),
        new TextRun({ text: `   ${co.ticker}   ·   ${co.currentPrice}`, font: "Arial", size: 22, bold: false, color: C.BLUE }),
      ],
      spacing: { before: 200, after: 120 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.NAVY, space: 1 } },
    }),
    sp(120),
  );

  children.push(verdictBadge(co.verdict));

  // Expected CAGR + prob of loss inline
  const lossColor = co.probOfLoss >= 25 ? C.RED : co.probOfLoss >= 12 ? C.AMBER : C.GREEN;
  const expColor2 = co.expectedCagr >= 14 ? C.GREEN : co.expectedCagr >= 8 ? C.AMBER : C.RED;
  children.push(
    paraRuns([
      { text: "Expected CAGR (probability-weighted): ", bold: true },
      { text: `${co.expectedCagr.toFixed(1)}%`, bold: true, color: expColor2 },
      { text: "   |   Probability of negative 10yr return: ", bold: true },
      { text: `${co.probOfLoss}%`, bold: true, color: lossColor },
    ])
  );
  children.push(sp(80));

  children.push(h3("Key Metrics"));
  children.push(metricsTable(co.metrics));
  children.push(sp(160));

  children.push(h3("Business Quality & Moat"));
  co.quality.forEach(q => children.push(para(q, { color: "222222" })));
  children.push(sp(80));

  children.push(h3("Disruption Risk"));
  children.push(riskPara(co.risk, co.riskText));
  children.push(sp(80));

  children.push(h3("4-Scenario Probability Analysis"));
  children.push(para("Disaster scenario reflects genuine business impairment — not just slow growth. Bear/Base/Bull reflect different growth trajectories. Probabilities reflect honest assessment of likelihood.", { italic: true, color: "666666" }));
  children.push(dcfTable(co.dcf, co.expectedCagr, co.probOfLoss));
  children.push(sp(160));

  children.push(h3("Verdict & Add Zone"));
  children.push(para(co.verdict_text, { color: "222222" }));
  children.push(paraRuns([
    { text: "Add Zone: ", bold: true },
    { text: co.addZone, bold: false, color: C.NAVY },
  ]));

  if (idx < companies.length - 1) children.push(pageBreak());
});

// ── PORTFOLIO PROBABILITY ANALYSIS ────────────────────────────────────────────
children.push(
  pageBreak(),
  h1("Portfolio-Level Probability Analysis"),
  divider(),

  h2("What the Numbers Honestly Show"),
  para("The probability-weighted expected CAGR across the ten positions, equally weighted, is approximately 11.0% per year. This is a genuinely good outcome — it implies roughly doubling the portfolio in real terms over ten years on a probability-weighted basis. But it is important to be clear about what the distribution of outcomes actually looks like."),

  h2("Probability of Achieving Key Return Thresholds"),
  sp(60),

  (() => {
    const W = [2200, 2600, 4000];
    const hdrs = ["Return Threshold", "Estimated Probability", "Commentary"];
    const hdrRow = new TableRow({
      children: hdrs.map((h, i) => new TableCell({
        borders, width: { size: W[i], type: WidthType.DXA },
        shading: { fill: C.NAVY, type: ShadingType.CLEAR }, margins: cellPad,
        children: [new Paragraph({ children: [new TextRun({ text: h, font: "Arial", size: 17, bold: true, color: C.WHITE })], alignment: AlignmentType.CENTER })],
      })),
    });
    const probRows = [
      ["Portfolio CAGR < 0%",  "~2–3%",   "Would require multiple simultaneous disaster scenarios. These companies are uncorrelated enough that this is very unlikely. A general market collapse followed by permanent impairment of several businesses is the only realistic path."],
      ["Portfolio CAGR 0–5%",  "~7–8%",   "Would require LVMH disaster + Copart AV disruption + one or two other disappointments simultaneously. Possible but unlikely given the diversity of business models and geographies."],
      ["Portfolio CAGR 5–10%", "~22–25%", "The 'decent but not exceptional' scenario. Most likely if LVMH and Copart both underperform, valuations stay stretched, and AI/tech growth decelerates. Still preserves and modestly grows capital in real terms."],
      ["Portfolio CAGR 10–15%","~40–45%", "The most probable single band. Represents base-case outcomes on most positions with one or two minor disappointments. This is roughly double the long-run index return from current valuations."],
      ["Portfolio CAGR > 15%", "~25–28%", "Requires the bull cases on the highest-conviction positions (CSU, Amazon, MELI) to play out AND no major disasters. Achievable but requires genuine compounding tailwinds."],
    ];
    const dataRows = probRows.map((r, idx) => {
      const fill = idx % 2 === 0 ? C.GRAY_BG : C.WHITE;
      return new TableRow({
        children: r.map((text, i) => new TableCell({
          borders, width: { size: W[i], type: WidthType.DXA },
          shading: { fill, type: ShadingType.CLEAR }, margins: cellPad,
          children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 17 })] })],
        })),
      });
    });
    return new Table({ width: { size: 8800, type: WidthType.DXA }, columnWidths: W, rows: [hdrRow, ...dataRows] });
  })(),

  sp(160),
  h2("The Single Honest Conclusion"),
  new Paragraph({
    children: [new TextRun({ text: "The probability of this portfolio delivering 5%+ over 10 years is approximately 90%. The probability of it delivering 10%+ is approximately 65–70%. The probability of it delivering 15%+ is approximately 25–28%.", font: "Arial", size: 20, bold: true, color: C.NAVY, italic: false })],
    spacing: { after: 160 },
    shading: { type: ShadingType.CLEAR, fill: C.LIGHT_BLUE },
    indent: { left: 200, right: 200 },
  }),
  sp(80),
  para("These are good odds for a buy-and-hold investor. The key risks that could move the portfolio toward the lower end are: (1) LVMH failing to recover and continuing to be a drag — at 37% loss probability, this is the portfolio's dead weight and should be reconsidered, (2) Copart experiencing accelerated AV disruption, and (3) a prolonged de-rating of quality multiples broadly. The key upside driver is CSU and Amazon compounding strongly over the full decade, which is the most likely scenario given their current positioning.", { color: "333333" }),

  h2("The LVMH Problem"),
  para("LVMH stands out starkly in the probability-weighted analysis. An expected CAGR of 3.7% with a 37% probability of negative returns is not consistent with the rest of the portfolio's quality standards. The business itself is excellent, but the valuation at €474 with declining revenues and significant structural uncertainty simply does not justify inclusion at current prices. This is the one position where the honest recommendation is: do not add, and consider reducing if the thesis does not show signs of recovery within 12–18 months.", { color: "333333" }),

  h2("The Strongest Positions by Probability-Weighted Analysis"),
  para("In order of expected CAGR: CSU/TOI (17.2%, 5% loss prob), Amazon (15.9%, 5% loss prob), MELI (14.1%, 12% loss prob), Microsoft (13.5%, 3% loss prob), Mastercard (10.7%, 8% loss prob). These five positions represent the portfolio's compounding engine. If they perform at their base-case expectations, they will more than compensate for any disappointments elsewhere. Overweighting these five relative to Ferrari, Copart, and LVMH would improve the portfolio's probability-weighted return.", { color: "333333" }),
);

// ── POSITION SIZING & DEPLOYMENT ─────────────────────────────────────────────
children.push(
  pageBreak(),
  h1("Position Sizing & Deployment Triggers"),
  divider(),

  para("This section translates the probability-weighted analysis into a concrete deployment plan. Position sizes are informed by expected CAGR and loss probability — higher-conviction positions (CSU, Amazon, MELI) deserve slightly larger initial allocations than lower-conviction ones (LVMH, Ferrari).", { color: "333333" }),

  h2("The Framework in Brief"),
  para("Total portfolio: ~£3m. Cash to deploy: £1.85m. Three existing positions already in place. The approach is three tranches per new position, sized to reach a maximum of 7–10% of total portfolio at cost (higher for GREEN positions, lower for RED). No more than two new positions opened per quarter."),
  para("The 3% / 3% / 2–3% tranche structure converts price weakness from a source of anxiety into a mechanical advantage: T2 and T3 trigger automatically on price drops, meaning you buy more at lower prices without requiring an emotional decision in the moment."),

  sp(80),
  h2("Existing Positions — Do You Add More?"),

  (() => {
    const W = [1700, 1000, 1000, 1400, 1900, 1760];
    const hdrs = ["Position", "Current Value", "% of Portfolio", "Cost % (est.)", "Add Trigger", "Rationale"];
    const hdrRow = new TableRow({
      children: hdrs.map((h, i) => new TableCell({
        borders, width: { size: W[i], type: WidthType.DXA },
        shading: { fill: C.NAVY, type: ShadingType.CLEAR }, margins: cellPad,
        children: [new Paragraph({ children: [new TextRun({ text: h, font: "Arial", size: 17, bold: true, color: C.WHITE })], alignment: AlignmentType.CENTER })],
      })),
    });
    const rows = [
      ["Games Workshop", "£340k", "11.3%", "~6.7% at cost", "Add only below £145", "Already above 8–9% target. Don't add at £170. Let it run. Watch £130–145 for a meaningful add."],
      ["MercadoLibre", "£250k", "8.3%", "~8.3% at cost*", "Small add below $1,500", "At target level. Could add £60–75k more if drops to $1,400–1,500. Otherwise position is full."],
      ["CSU / Topicus", "£400k", "13.3%", "~14.8% at cost*", "Add only below $2,300 CAD", "Well over target due to original sizing. Let it compound. Only add on a 10%+ pullback."],
    ];
    const dataRows = rows.map((r, idx) => {
      const fill = idx % 2 === 0 ? C.GRAY_BG : C.WHITE;
      return new TableRow({
        children: r.map((text, i) => new TableCell({
          borders, width: { size: W[i], type: WidthType.DXA },
          shading: { fill, type: ShadingType.CLEAR }, margins: cellPad,
          children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 17 })] })],
        })),
      });
    });
    return new Table({ width: { size: 8760, type: WidthType.DXA }, columnWidths: W, rows: [hdrRow, ...dataRows] });
  })(),
  sp(80),
  para("* CSU/TOI cost basis estimated from reported entry prices ($2,800 CAD / $110 CAD) vs current. MELI cost basis estimated from $1,920 entry. These existing positions are already at or above target allocation. The priority for the £1.85m cash is the seven new positions below.", { italic: true, color: "666666" }),

  sp(160),
  h2("New Positions — Full Deployment Plan"),
  para("Position sizing reflects the probability-weighted analysis. GREEN positions receive larger initial tranches (3%). AMBER positions receive 1.7–2.5% initially. RED positions have no immediate buy — waiting for valuation improvement."),
  sp(80),

  (() => {
    const W = [1400, 700, 1600, 1800, 1800, 1460];
    const hdrs = ["Company", "Signal", "T1 — Buy Now", "T2 — Price OR Time Trigger", "T3 — Price OR Time Trigger", "Max Position"];
    const hdrRow = new TableRow({
      children: hdrs.map((h, i) => new TableCell({
        borders, width: { size: W[i], type: WidthType.DXA },
        shading: { fill: C.NAVY, type: ShadingType.CLEAR }, margins: cellPad,
        children: [new Paragraph({ children: [new TextRun({ text: h, font: "Arial", size: 16, bold: true, color: C.WHITE })], alignment: AlignmentType.CENTER })],
      })),
    });
    const positions = [
      { name: "Amazon",     signal: "GREEN", t1: "3% / ~£90k\nat ~$210\nBuy now",             t2: "3% / ~£90k\nAt $175 OR 6 months",          t3: "2.5% / ~£75k\nAt $150 OR 12 months",          max: "8.5% / ~£255k", signalColor: C.GREEN_BG },
      { name: "Copart",     signal: "AMBER", t1: "2.5% / ~£75k\nat ~$33\nBuy now",            t2: "3% / ~£90k\nAt $27–28 OR 6 months",         t3: "2.5% / ~£75k\nAt $23–24 OR 12 months",         max: "8% / ~£240k",   signalColor: C.AMBER_BG },
      { name: "S&P Global", signal: "AMBER", t1: "1.7% / ~£50k\nat ~$426\nSmall T1 only",     t2: "3% / ~£90k\nAt $390–400 OR 6 months",       t3: "2.5% / ~£75k\nAt $355–370 OR 12 months",       max: "7% / ~£215k",   signalColor: C.AMBER_BG },
      { name: "Mastercard", signal: "AMBER", t1: "1.7% / ~£50k\nat ~$507\nSmall T1 only",     t2: "3% / ~£90k\nAt $415–425 OR 6 months",       t3: "2.5% / ~£75k\nAt $375–390 OR 12 months",       max: "7% / ~£215k",   signalColor: C.AMBER_BG },
      { name: "Microsoft",  signal: "AMBER", t1: "2% / ~£60k\nat ~$399\nSmall T1 only",       t2: "3% / ~£90k\nAt $360–370 OR 6 months",       t3: "2.5% / ~£75k\nAt $320–340 OR 12 months",       max: "7.5% / ~£225k", signalColor: C.AMBER_BG },
      { name: "Ferrari",    signal: "RED",   t1: "0% — Wait\nDo not buy yet",                 t2: "2.5% / ~£75k\nOnly at $255–260",             t3: "3% / ~£90k\nOnly at $220–235",                 max: "5.5% / ~£165k\n(lower max — RED)", signalColor: C.RED_BG },
      { name: "LVMH",       signal: "RED",   t1: "0% — Wait\nDo not buy yet",                 t2: "2.5% / ~£75k\nOnly at €350–360",             t3: "3% / ~£90k\nOnly at €295–310",                 max: "5.5% / ~£165k\n(lower max — RED)", signalColor: C.RED_BG },
    ];
    const signalTextColor = { GREEN: C.GREEN, AMBER: C.AMBER, RED: C.RED };
    const dataRows = positions.map((p, idx) => {
      const fill = idx % 2 === 0 ? C.GRAY_BG : C.WHITE;
      return new TableRow({
        children: [
          new TableCell({ borders, width: { size: W[0], type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: cellPad, children: [new Paragraph({ children: [new TextRun({ text: p.name, font: "Arial", size: 17, bold: true })] })] }),
          new TableCell({ borders, width: { size: W[1], type: WidthType.DXA }, shading: { fill: p.signalColor, type: ShadingType.CLEAR }, margins: cellPad, children: [new Paragraph({ children: [new TextRun({ text: p.signal, font: "Arial", size: 16, bold: true, color: signalTextColor[p.signal] })], alignment: AlignmentType.CENTER })] }),
          new TableCell({ borders, width: { size: W[2], type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: cellPad, children: p.t1.split("\n").map(line => new Paragraph({ children: [new TextRun({ text: line, font: "Arial", size: 16 })] })) }),
          new TableCell({ borders, width: { size: W[3], type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: cellPad, children: p.t2.split("\n").map(line => new Paragraph({ children: [new TextRun({ text: line, font: "Arial", size: 16 })] })) }),
          new TableCell({ borders, width: { size: W[4], type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: cellPad, children: p.t3.split("\n").map(line => new Paragraph({ children: [new TextRun({ text: line, font: "Arial", size: 16 })] })) }),
          new TableCell({ borders, width: { size: W[5], type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: cellPad, children: p.max.split("\n").map(line => new Paragraph({ children: [new TextRun({ text: line, font: "Arial", size: 16, bold: true })] })) }),
        ],
      });
    });
    return new Table({ width: { size: 8760, type: WidthType.DXA }, columnWidths: W, rows: [hdrRow, ...dataRows] });
  })(),

  sp(80),
  para("IMPORTANT: T2 and T3 are triggered by WHICHEVER COMES FIRST — the price target OR the time elapsed. If a position is up 20% six months after T1, the time trigger still fires. You re-evaluate the thesis, confirm it is intact, and if so you add Tranche 2 regardless of price. This is the key rule that prevents the 'up 20% so I won't add' paralysis.", { color: "444444" }),

  sp(160),
  h2("Deployment Cash Flow Over Time"),
  para("Assuming no major market move, here is approximately how the £1.85m deploys:"),
  sp(60),

  (() => {
    const W = [1200, 2800, 2000, 2760];
    const hdrs = ["Period", "Action", "Cash Deployed", "Remaining Cash"];
    const hdrRow = new TableRow({
      children: hdrs.map((h, i) => new TableCell({
        borders, width: { size: W[i], type: WidthType.DXA },
        shading: { fill: C.NAVY, type: ShadingType.CLEAR }, margins: cellPad,
        children: [new Paragraph({ children: [new TextRun({ text: h, font: "Arial", size: 17, bold: true, color: C.WHITE })], alignment: AlignmentType.CENTER })],
      })),
    });
    const cashRows = [
      ["Now (Q1)", "T1: Amazon £90k, Copart £75k, SPGI £50k, Mastercard £50k, Microsoft £60k", "~£325k", "~£1,525k"],
      ["Q2 (3 months)", "T2s fire if prices fall. If no falls: MELI add £60k (if below $1,500). CSU add £75k (if below $2,300 CAD).", "£0–£135k", "~£1,390–1,525k"],
      ["Q3 (6 months)", "Time triggers: T2s on Amazon, Copart, SPGI, MA, MSFT (if not already triggered by price). ~£450k total if all fire.", "£200–£450k", "~£940–1,190k"],
      ["Q4 (9 months)", "T3s begin firing. Ferrari T2 if at $255–260. Further T2s on existing names.", "£150–£350k", "~£590–1,040k"],
      ["Q5–Q6 (12–18 months)", "Remaining T3s complete. LVMH T2 if at €350. Full portfolio built out on 7 new positions.", "£300–£500k", "~£90–740k"],
      ["Q7–Q8 (18–24 months)", "Final tranches + any opportunistic adds. Target: fully deployed, holding ~£150–200k crash reserve.", "Remaining", "£150–200k reserve"],
    ];
    const dataRows = cashRows.map((r, idx) => {
      const fill = idx % 2 === 0 ? C.GRAY_BG : C.WHITE;
      return new TableRow({
        children: r.map((text, i) => new TableCell({
          borders, width: { size: W[i], type: WidthType.DXA },
          shading: { fill, type: ShadingType.CLEAR }, margins: cellPad,
          children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 16 })] })],
        })),
      });
    });
    return new Table({ width: { size: 8760, type: WidthType.DXA }, columnWidths: W, rows: [hdrRow, ...dataRows] });
  })(),

  sp(160),
  h2("The Bear Market Scenario — Your Most Important Pre-Commitment"),
  para("A market correction is not a risk to this strategy — it is the strategy's greatest ally, if you pre-commit to responding correctly. Quality businesses fall in corrections but recover and compound past their pre-correction levels. The investor who adds at -25% versus the one who freezes ends up in dramatically different places after 10 years."),
  sp(80),

  (() => {
    const W = [2000, 2600, 4160];
    const hdrs = ["Scenario", "Portfolio Impact", "What You Do (Pre-Committed)"];
    const hdrRow = new TableRow({
      children: hdrs.map((h, i) => new TableCell({
        borders, width: { size: W[i], type: WidthType.DXA },
        shading: { fill: C.NAVY, type: ShadingType.CLEAR }, margins: cellPad,
        children: [new Paragraph({ children: [new TextRun({ text: h, font: "Arial", size: 17, bold: true, color: C.WHITE })], alignment: AlignmentType.CENTER })],
      })),
    });
    const scenarioRows = [
      ["Market falls 10–15%\n(routine correction)", "Deployed positions fall £80–130k on paper. Cash still ~£1.4–1.5m.", "Continue normal deployment pace. Time triggers still fire quarterly. Do not slow down. This is noise."],
      ["Market falls 20–25%\n(mild bear market)", "Deployed positions fall £150–200k on paper. Most T2 price triggers fire automatically.", "ACCELERATE: deploy all pending T2 tranches immediately, ahead of schedule. This is exactly what you want — buying quality at 20% below your original entry."],
      ["Market falls 30–35%\n(severe bear market)", "Deployed positions fall £200–280k on paper. Psychologically very hard.", "MAXIMUM AGGRESSION: deploy all remaining T2 and T3 tranches immediately. Consider increasing position sizes to 10–12% on highest-conviction names (Amazon, MELI, CSU). This scenario, while painful, is the single best buying opportunity you will likely see."],
      ["Market flat for 2–3 years\n(grinding sideways)", "Portfolio grows slowly. Cash drag from uninvested pounds.", "Stick to the time triggers. A flat market means you deploy T2s and T3s at the same or similar prices to T1 — not ideal but perfectly acceptable for 10-year quality holds. Do not abandon the plan out of impatience."],
      ["Market rises 20–30%\n(continues running)", "Deployed positions up significantly. Price triggers not hit, time triggers fire.", "Allow time triggers to fire quarterly regardless of price. Accept that T2 and T3 are at higher prices than T1. The businesses are compounding. Don't let price appreciation stop you completing the position."],
    ];
    const dataRows = scenarioRows.map((r, idx) => {
      const fill = idx % 2 === 0 ? C.GRAY_BG : C.WHITE;
      return new TableRow({
        children: r.map((text, i) => new TableCell({
          borders, width: { size: W[i], type: WidthType.DXA },
          shading: { fill, type: ShadingType.CLEAR }, margins: cellPad,
          children: text.split("\n").map(line => new Paragraph({ children: [new TextRun({ text: line, font: "Arial", size: 16 })] })),
        })),
      });
    });
    return new Table({ width: { size: 8760, type: WidthType.DXA }, columnWidths: W, rows: [hdrRow, ...dataRows] });
  })(),

  sp(160),
  h2("The Cash Reserve Rule"),
  para("Keep approximately £150–200k in cash permanently — even after full deployment — as a crash reserve. This only deploys when a single position falls 25%+ from your cost basis AND you can confirm the thesis is intact. Otherwise it sits idle. It is not a market-timing reserve — it is pre-allocated firepower for a sudden, severe, thesis-intact sell-off in one of your highest-conviction positions."),

  sp(160),
  h2("The Single Most Important Rule"),
  new Paragraph({
    children: [new TextRun({ text: "Write down your thesis for each position before you buy Tranche 1. Include: what makes this business exceptional, what would make you sell (specific business deterioration, not price falling), and your T2 and T3 triggers. When you feel the urge to deviate — either to buy more than planned or to sell during a drawdown — read what you wrote when you were calm. The plan you make when unemotional is almost always better than the decision you make when staring at a red portfolio.", font: "Arial", size: 20, bold: false, color: C.NAVY, italic: true })],
    spacing: { after: 200 },
    shading: { type: ShadingType.CLEAR, fill: C.LIGHT_BLUE },
  }),

  sp(400),
  new Paragraph({
    children: [new TextRun({ text: "This analysis is for informational purposes only and does not constitute financial advice. Past returns do not guarantee future results. All estimates are forward-looking and inherently uncertain.", font: "Arial", size: 16, italic: true, color: "888888" })],
    alignment: AlignmentType.CENTER,
  }),
);

// ─── Build & write ────────────────────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: C.NAVY },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: C.BLUE },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: C.NAVY },
        paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } },
    ],
  },
  sections: [{
    properties: {
      page: {
        margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          children: [
            new TextRun({ text: "Concentrated Portfolio Analysis — March 2026 — PROBABILITY-WEIGHTED FRAMEWORK", font: "Arial", size: 16, color: "888888" }),
          ],
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: C.BORDER } },
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          children: [
            new TextRun({ text: "For informational purposes only. Not financial advice.  |  Page ", font: "Arial", size: 16, color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: "888888" }),
          ],
          alignment: AlignmentType.CENTER,
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync('/sessions/trusting-compassionate-thompson/mnt/outputs/Portfolio_Analysis_March2026.docx', buffer);
  console.log('Document written successfully.');
});
