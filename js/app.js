/**
 * SEC Financial & Sentiment Analyzer — Interactive Dashboard Application
 * Authoritative data sources: docs/data/*.json & Python SEC REST API Backend
 */

// Application State
const state = {
  currentCompany: 'AAPL', // 'AAPL', 'AMD', 'COMP', or 'CUSTOM'
  currentTab: 'overview',
  currentStatement: 'is', // 'is', 'bs', 'cf'
  chartInstance: null,
  compChartInstance: null,
  ratioChartInstance: null,
  multiCompChartInstance: null,
  customCompany: null, // Microsoft Corporation (MSFT) analysis object
  multiComp: {
    enabledCompanies: ['AAPL', 'AMD', 'CUSTOM'],
    selectedMetric: 'net_sales_m',
    startYear: 'FY2016',
    endYear: 'FY2025',
    chartType: 'line' // 'line' or 'bar'
  },
  pairComp: {
    selectedMetric: 'net_sales_m',
    startYear: 'FY2016',
    endYear: 'FY2025',
    chartType: 'line',
    chartInstance: null,
    trajectoryChartInstance: null
  },
  data: {
    comparison: null,
    aapl: null,
    amd: null,
    msft_precomputed: null
  }
};

// Company Profiles
const PROFILES = {
  AAPL: {
    name: 'Apple Inc.',
    ticker: 'AAPL',
    cik: '0000320193',
    sector: 'Technology',
    industry: 'Consumer Electronics',
    fiscalEnd: 'Last Saturday of September',
    exchange: 'NASDAQ',
    color: '#38bdf8',
    description: 'Global designer and manufacturer of smartphones, personal computers, tablets, wearables, and related digital services.'
  },
  AMD: {
    name: 'Advanced Micro Devices, Inc.',
    ticker: 'AMD',
    cik: '0000002488',
    sector: 'Technology',
    industry: 'Semiconductors',
    fiscalEnd: 'Last Saturday of December',
    exchange: 'NASDAQ',
    color: '#ef4444',
    description: 'Global semiconductor company developing high-performance computing, graphics, and data center visualization technologies.'
  },
  CUSTOM: {
    name: 'Microsoft Corp.',
    ticker: 'MSFT',
    cik: '0000789019',
    sector: 'Technology',
    industry: 'Systems & Enterprise Software / Cloud',
    fiscalEnd: 'June 30',
    exchange: 'NASDAQ',
    color: '#10b981',
    description: 'Global provider of enterprise cloud computing (Azure), productivity software (Office 365), and personal computing solutions. Peer financial benchmark comparison company.'
  }
};

// Metric Definitions & Accounting Benchmarks
const METRIC_DEFINITIONS = {
  net_sales_m: {
    name: 'Total Net Sales / Revenue ($M)',
    category: 'Top-Line Financial Performance',
    definition: 'Total operating revenue recognized from goods, services, and product sales after returns and discounts.',
    unit: '$M',
    format: (v) => formatMillions(v),
    interpret: 'Indicates gross commercial scale and aggregate market demand across operating segments.'
  },
  net_income_m: {
    name: 'Net Income ($M)',
    category: 'Bottom-Line Financial Performance',
    definition: 'Total accounting profit remaining after all operating expenses, cost of sales, depreciation, interest, and income taxes.',
    unit: '$M',
    format: (v) => formatMillions(v),
    interpret: 'Fundamental metric of bottom-line equity value generation; supports negative values during corporate turnarounds.'
  },
  revenue_growth_pct: {
    name: 'Revenue Growth YoY (%)',
    category: 'Operating Velocity',
    definition: 'Year-over-year percentage change in total net sales: (Revenue_t - Revenue_{t-1}) / Revenue_{t-1}.',
    unit: '%',
    format: (v) => formatPercent(v),
    interpret: 'Measures organic and acquisitive sales expansion speed.'
  },
  net_income_growth_pct: {
    name: 'Net Income Growth YoY (%)',
    category: 'Profitability Velocity',
    definition: 'Year-over-year percentage change in net income: (NI_t - NI_{t-1}) / |NI_{t-1}|.',
    unit: '%',
    format: (v) => formatPercent(v),
    interpret: 'Illustrates operating leverage and earnings sensitivity.'
  },
  total_assets_m: {
    name: 'Total Assets ($M)',
    category: 'Capital Base',
    definition: 'Total economic resources owned and controlled by the company reported on the balance sheet at fiscal year-end.',
    unit: '$M',
    format: (v) => formatMillions(v),
    interpret: 'Measures total productive capital deployed in the enterprise.'
  },
  operating_cf_m: {
    name: 'Cash from Operations ($M)',
    category: 'Cash Flow Generation',
    definition: 'Net cash provided by operating activities from the Statement of Cash Flows.',
    unit: '$M',
    format: (v) => formatMillions(v),
    interpret: 'True cash earnings generated from core operational activities, excluding non-cash accounting accruals.'
  },
  net_profit_margin: {
    name: 'Net Profit Margin (%) [Profitability]',
    category: 'Profitability Ratio',
    definition: 'Net Income divided by Total Net Sales (Net Income / Revenue).',
    benchmark: 'Industry Benchmark: > 15.0%',
    unit: '%',
    format: (v) => formatPercent(v),
    interpret: 'Measures bottom-line operational conversion efficiency and structural pricing power.'
  },
  current_ratio: {
    name: 'Current Ratio (x) [Liquidity]',
    category: 'Liquidity Ratio',
    definition: 'Total Current Assets divided by Total Current Liabilities.',
    benchmark: 'Standard Benchmark: 1.00x - 2.00x',
    unit: 'x',
    format: (v) => formatRatio(v),
    interpret: 'Evaluates short-term solvency buffer and ability to cover near-term obligations due within 12 months.'
  },
  liabilities_to_assets: {
    name: 'Liabilities-to-Assets Ratio (%) [Solvency]',
    category: 'Solvency / Financial Leverage',
    definition: 'Total Liabilities divided by Total Assets.',
    benchmark: 'Prudent Benchmark: < 60.0%',
    unit: '%',
    format: (v) => formatPercent(v),
    interpret: 'Reflects capital structure leverage, debt exposure, and financial risk relative to total enterprise resources.'
  },
  asset_turnover: {
    name: 'Total Asset Turnover (x) [Efficiency]',
    category: 'Efficiency / Asset Velocity',
    definition: 'Total Net Sales divided by Average Total Assets ((Assets_t + Assets_{t-1}) / 2).',
    benchmark: 'Efficiency Benchmark: > 0.50x',
    unit: 'x',
    format: (v) => formatRatio(v),
    interpret: 'Measures revenue productivity generated per dollar of capital asset deployment.'
  },
  pn_ratio: {
    name: 'MD&A Positive / Negative Ratio (x)',
    category: 'Linguistic Sentiment',
    definition: 'Total matched Loughran-McDonald positive words divided by negative words in Form 10-K Item 7 narrative.',
    benchmark: 'Neutral Benchmark: 1.00x (> 1.0 = Net Optimistic; < 1.0 = Net Cautious)',
    unit: 'x',
    format: (v) => (v !== null && v !== undefined ? Number(v).toFixed(4) + 'x' : '—'),
    interpret: 'Domain-specific accounting sentiment ratio extracted from annual management disclosures.'
  },
  net_sentiment_k: {
    name: 'Net Sentiment per 1,000 Words',
    category: 'Linguistic Sentiment Density',
    definition: '((Positive Words - Negative Words) / Total MD&A Word Count) * 1,000.',
    unit: 'per 1k',
    format: (v) => (v !== null && v !== undefined ? (Number(v) >= 0 ? '+' : '') + Number(v).toFixed(2) : '—'),
    interpret: 'Normalized net narrative optimism density across annual management disclosures.'
  },
  total_word_count: {
    name: 'Total MD&A Word Count',
    category: 'Disclosure Length',
    definition: 'Aggregate word token length of Item 7 narrative in primary Form 10-K filing.',
    unit: 'words',
    format: (v) => formatNumber(v),
    interpret: 'Reflects length, complexity, and thoroughness of management discussion and risk disclosure.'
  },
  pos_words: {
    name: 'Positive Word Count',
    category: 'Linguistic Sentiment',
    definition: 'Frequency count of words matching the Loughran-McDonald Positive category in Item 7.',
    unit: 'words',
    format: (v) => formatNumber(v),
    interpret: 'Optimistic narrative signals.'
  },
  neg_words: {
    name: 'Negative Word Count',
    category: 'Linguistic Sentiment',
    definition: 'Frequency count of words matching the Loughran-McDonald Negative category in Item 7.',
    unit: 'words',
    format: (v) => formatNumber(v),
    interpret: 'Cautionary, risk, and adversity-oriented narrative signals.'
  },
  split_adj_close: {
    name: 'Split-Adjusted Closing Stock Price ($)',
    category: 'Market Valuation',
    definition: 'Market closing equity price at fiscal year-end, rolled back to preceding trading day for weekend/holiday dates.',
    unit: '$',
    format: (v) => formatCurrency(v),
    interpret: 'Historical market closing price adjusted for forward and reverse stock splits.'
  },
  indexed_price: {
    name: 'Normalized Stock Price (FY2016 = 100.0)',
    category: 'Market Valuation',
    definition: 'Split-adjusted stock price indexed to baseline FY2016 = 100.0 for direct comparative equity return analysis.',
    unit: 'Index',
    format: (v) => (v !== null && v !== undefined ? Number(v).toFixed(1) : '—'),
    interpret: 'Relative equity price appreciation over the study period, removing initial share price magnitude bias.'
  }
};

// Formatting Utilities
const formatCurrency = (val) => {
  if (val === null || val === undefined || isNaN(val)) return '—';
  return '$' + Number(val).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};

const formatMillions = (val) => {
  if (val === null || val === undefined || isNaN(val)) return '—';
  return '$' + Number(val).toLocaleString('en-US', { maximumFractionDigits: 0 }) + 'M';
};

const formatPercent = (val) => {
  if (val === null || val === undefined || isNaN(val)) return '—';
  const pct = Number(val) * 100;
  return (pct >= 0 ? '+' : '') + pct.toFixed(1) + '%';
};

const formatRatio = (val) => {
  if (val === null || val === undefined || isNaN(val)) return '—';
  return Number(val).toFixed(4) + 'x';
};

const formatNumber = (val) => {
  if (val === null || val === undefined || isNaN(val)) return '—';
  return Number(val).toLocaleString('en-US');
};

// Initialize Application
async function initApp() {
  try {
    if (window.SEC_DATA && window.SEC_DATA.comparison) {
      state.data.comparison = window.SEC_DATA.comparison;
      state.data.aapl = window.SEC_DATA.aapl;
      state.data.amd = window.SEC_DATA.amd;
      state.data.msft_precomputed = window.SEC_DATA.msft_precomputed || null;
      console.log('[SEC Analyzer] Authoritative datasets loaded from data_bundle.js (zero latency / offline compatible).');
    } else {
      // Fetch authoritative JSON files
      const [compRes, aaplRes, amdRes] = await Promise.all([
        fetch('data/comparison_data.json'),
        fetch('data/aapl_data.json'),
        fetch('data/amd_data.json')
      ]);

      state.data.comparison = await compRes.json();
      state.data.aapl = await aaplRes.json();
      state.data.amd = await amdRes.json();

      try {
        const msftRes = await fetch('data/msft_precomputed.json');
        if (msftRes.ok) {
          state.data.msft_precomputed = await msftRes.json();
        }
      } catch (e) {
        console.log('No precomputed MSFT file present.');
      }

      console.log('[SEC Analyzer] Authoritative datasets successfully loaded via fetch.');
    }
    
    // Preload verified Microsoft dataset
    if (state.data.msft_precomputed) {
      state.customCompany = state.data.msft_precomputed;
      if (!state.multiComp.enabledCompanies.includes('CUSTOM')) {
        state.multiComp.enabledCompanies.push('CUSTOM');
      }
    }
    
    populateYearDropdowns();
    setupEventListeners();
    renderAll();
  } catch (err) {
    console.error('[SEC Analyzer] Failed to load data JSON files:', err);
    document.getElementById('app-loading').innerHTML = `
      <div class="callout disclaimer" style="margin: 2rem;">
        <h4>Failed to Load Data</h4>
        <p>Could not load precomputed JSON data. Please ensure the website is accessed via an HTTP/HTTPS server (e.g. GitHub Pages or local web server) rather than direct file:// access.</p>
        <p style="font-family: monospace; margin-top: 0.5rem;">Error: ${err.message}</p>
      </div>
    `;
  }
}

// Populate Comparison Year Range Dropdowns
function populateYearDropdowns() {
  const years = state.data.comparison.fiscal_years;

  // Multi-Company Dropdowns
  const selStart = document.getElementById('select-year-start');
  const selEnd = document.getElementById('select-year-end');
  if (selStart && selEnd) {
    selStart.innerHTML = '';
    selEnd.innerHTML = '';
    years.forEach((yr, idx) => {
      const opt1 = document.createElement('option');
      opt1.value = yr;
      opt1.textContent = yr;
      if (idx === 0) opt1.selected = true;
      selStart.appendChild(opt1);

      const opt2 = document.createElement('option');
      opt2.value = yr;
      opt2.textContent = yr;
      if (idx === years.length - 1) opt2.selected = true;
      selEnd.appendChild(opt2);
    });
  }

  // Pair Comparison Dropdowns (Apple vs AMD)
  const selPairStart = document.getElementById('select-pair-year-start');
  const selPairEnd = document.getElementById('select-pair-year-end');
  if (selPairStart && selPairEnd) {
    selPairStart.innerHTML = '';
    selPairEnd.innerHTML = '';
    years.forEach((yr, idx) => {
      const opt1 = document.createElement('option');
      opt1.value = yr;
      opt1.textContent = yr;
      if (idx === 0) opt1.selected = true;
      selPairStart.appendChild(opt1);

      const opt2 = document.createElement('option');
      opt2.value = yr;
      opt2.textContent = yr;
      if (idx === years.length - 1) opt2.selected = true;
      selPairEnd.appendChild(opt2);
    });
  }

  state.multiComp.startYear = years[0];
  state.multiComp.endYear = years[years.length - 1];
  state.pairComp.startYear = years[0];
  state.pairComp.endYear = years[years.length - 1];
}

// Setup Event Listeners
function setupEventListeners() {
  // Company Switcher Buttons
  document.querySelectorAll('.company-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      // If clicking remove cross, handle removal
      if (e.target.classList.contains('btn-remove-company')) {
        e.stopPropagation();
        removeCustomCompany();
        return;
      }

      const target = e.currentTarget.dataset.company;
      if (target === state.currentCompany && target !== 'COMP') return;

      document.querySelectorAll('.company-btn').forEach(b => b.classList.remove('active'));
      e.currentTarget.classList.add('active');

      state.currentCompany = target;

      if (target === 'COMP') {
        switchTab('cross-company');
      }
      renderAll();
    });
  });

  // Navigation Links
  document.querySelectorAll('.nav-link').forEach(link => {
    link.addEventListener('click', (e) => {
      e.preventDefault();
      const tabId = e.currentTarget.dataset.tab;
      switchTab(tabId);
    });
  });

  // Quick Analyze Button in Header
  const btnNavAdd = document.getElementById('btn-nav-add-company');
  if (btnNavAdd) {
    btnNavAdd.addEventListener('click', () => {
      switchTab('add-company');
    });
  }

  // Financial Statement Toggle Buttons
  document.querySelectorAll('.btn-statement').forEach(btn => {
    btn.addEventListener('click', (e) => {
      document.querySelectorAll('.btn-statement').forEach(b => b.classList.remove('active'));
      e.currentTarget.classList.add('active');
      state.currentStatement = e.currentTarget.dataset.statement;
      renderFinancialStatements();
    });
  });

  // Third Company Analysis Trigger
  const btnRunAnalysis = document.getElementById('btn-run-analysis');
  const inputTicker = document.getElementById('input-third-ticker');

  if (btnRunAnalysis && inputTicker) {
    btnRunAnalysis.addEventListener('click', () => {
      const ticker = inputTicker.value.trim().toUpperCase();
      if (!ticker) {
        showAnalysisAlert('Please enter a stock ticker symbol (e.g., MSFT, NVDA, TSLA).', 'warning');
        return;
      }
      executeTickerAnalysis(ticker);
    });

    inputTicker.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        btnRunAnalysis.click();
      }
    });
  }

  // Quick Chips
  document.querySelectorAll('.quick-chip').forEach(chip => {
    chip.addEventListener('click', (e) => {
      const ticker = e.currentTarget.dataset.quick;
      if (inputTicker) {
        inputTicker.value = ticker;
        executeTickerAnalysis(ticker);
      }
    });
  });

  // Instant Precomputed MSFT Baseline Button
  const btnLoadPrecomputed = document.getElementById('btn-load-precomputed-msft');
  if (btnLoadPrecomputed) {
    btnLoadPrecomputed.addEventListener('click', () => {
      loadPrecomputedMSFT();
    });
  }

  // Jump Buttons in Analysis Success Card & Peer Benchmark Card
  const btnJumpView = document.getElementById('btn-jump-view-company');
  if (btnJumpView) {
    btnJumpView.addEventListener('click', () => {
      if (state.customCompany) {
        state.currentCompany = 'CUSTOM';
        document.querySelectorAll('.company-btn').forEach(b => b.classList.remove('active'));
        const customBtn = document.getElementById('btn-custom-company');
        if (customBtn) customBtn.classList.add('active');
        switchTab('overview');
        renderAll();
      }
    });
  }

  const btnJumpComp = document.getElementById('btn-jump-compare');
  if (btnJumpComp) {
    btnJumpComp.addEventListener('click', () => {
      state.currentCompany = 'COMP';
      document.querySelectorAll('.company-btn').forEach(b => b.classList.remove('active'));
      const compBtn = document.getElementById('btn-comp-mode');
      if (compBtn) compBtn.classList.add('active');
      switchTab('cross-company');
      renderAll();
    });
  }

  const btnJumpViewMsft = document.getElementById('btn-jump-view-msft');
  if (btnJumpViewMsft) {
    btnJumpViewMsft.addEventListener('click', () => {
      state.currentCompany = 'CUSTOM';
      document.querySelectorAll('.company-btn').forEach(b => b.classList.remove('active'));
      const customBtn = document.getElementById('btn-custom-company');
      if (customBtn) customBtn.classList.add('active');
      switchTab('financials');
      renderAll();
    });
  }

  const btnJumpCompMsft = document.getElementById('btn-jump-compare-msft');
  if (btnJumpCompMsft) {
    btnJumpCompMsft.addEventListener('click', () => {
      state.currentCompany = 'COMP';
      document.querySelectorAll('.company-btn').forEach(b => b.classList.remove('active'));
      const compBtn = document.getElementById('btn-comp-mode');
      if (compBtn) compBtn.classList.add('active');
      switchTab('cross-company');
      renderAll();
    });
  }

  // Multi-Company Toolbar Controls
  const chkAapl = document.getElementById('chk-comp-aapl');
  const chkAmd = document.getElementById('chk-comp-amd');
  const chkCustom = document.getElementById('chk-comp-custom');

  [chkAapl, chkAmd, chkCustom].forEach(chk => {
    if (chk) {
      chk.addEventListener('change', () => {
        updateEnabledCompanies();
        renderMultiCompanySection();
      });
    }
  });

  // Remove Third Company Buttons
  const btnRemoveToolbar = document.getElementById('btn-remove-company-toolbar');
  if (btnRemoveToolbar) {
    btnRemoveToolbar.addEventListener('click', (e) => {
      e.stopPropagation();
      removeCustomCompany();
    });
  }

  const selectMetric = document.getElementById('select-comp-metric');
  if (selectMetric) {
    selectMetric.addEventListener('change', (e) => {
      state.multiComp.selectedMetric = e.target.value;
      renderMultiCompanySection();
    });
  }

  const selStart = document.getElementById('select-year-start');
  const selEnd = document.getElementById('select-year-end');
  if (selStart && selEnd) {
    selStart.addEventListener('change', (e) => {
      state.multiComp.startYear = e.target.value;
      renderMultiCompanySection();
    });
    selEnd.addEventListener('change', (e) => {
      state.multiComp.endYear = e.target.value;
      renderMultiCompanySection();
    });
  }

  // Chart Type Toggle
  const btnLine = document.getElementById('btn-chart-type-line');
  const btnBar = document.getElementById('btn-chart-type-bar');
  if (btnLine && btnBar) {
    btnLine.addEventListener('click', () => {
      btnLine.classList.add('active');
      btnBar.classList.remove('active');
      state.multiComp.chartType = 'line';
      renderMultiCompanySection();
    });
    btnBar.addEventListener('click', () => {
      btnBar.classList.add('active');
      btnLine.classList.remove('active');
      state.multiComp.chartType = 'bar';
      renderMultiCompanySection();
    });
  }

  // Export Buttons
  const btnExportPng = document.getElementById('btn-export-chart-png');
  if (btnExportPng) {
    btnExportPng.addEventListener('click', exportChartPNG);
  }

  const btnExportCsv = document.getElementById('btn-export-comp-csv');
  if (btnExportCsv) {
    btnExportCsv.addEventListener('click', exportComparisonCSV);
  }

  // Pair Comparison (Apple vs AMD) Toolbar Controls
  const selectPairMetric = document.getElementById('select-pair-metric');
  if (selectPairMetric) {
    selectPairMetric.addEventListener('change', (e) => {
      state.pairComp.selectedMetric = e.target.value;
      renderPairSection();
    });
  }

  const selPairStart = document.getElementById('select-pair-year-start');
  const selPairEnd = document.getElementById('select-pair-year-end');
  if (selPairStart && selPairEnd) {
    selPairStart.addEventListener('change', (e) => {
      state.pairComp.startYear = e.target.value;
      renderPairSection();
    });
    selPairEnd.addEventListener('change', (e) => {
      state.pairComp.endYear = e.target.value;
      renderPairSection();
    });
  }

  const btnPairLine = document.getElementById('btn-pair-type-line');
  const btnPairBar = document.getElementById('btn-pair-type-bar');
  if (btnPairLine && btnPairBar) {
    btnPairLine.addEventListener('click', () => {
      btnPairLine.classList.add('active');
      btnPairBar.classList.remove('active');
      state.pairComp.chartType = 'line';
      renderPairSection();
    });
    btnPairBar.addEventListener('click', () => {
      btnPairBar.classList.add('active');
      btnPairLine.classList.remove('active');
      state.pairComp.chartType = 'bar';
      renderPairSection();
    });
  }

  const btnExportPairPng = document.getElementById('btn-export-pair-png');
  if (btnExportPairPng) {
    btnExportPairPng.addEventListener('click', exportPairPNG);
  }

  const btnExportPairCsv = document.getElementById('btn-export-pair-csv');
  if (btnExportPairCsv) {
    btnExportPairCsv.addEventListener('click', exportPairCSV);
  }
}

function updateEnabledCompanies() {
  const enabled = [];
  if (document.getElementById('chk-comp-aapl')?.checked) enabled.push('AAPL');
  if (document.getElementById('chk-comp-amd')?.checked) enabled.push('AMD');
  if (state.customCompany && document.getElementById('chk-comp-custom')?.checked) enabled.push('CUSTOM');
  state.multiComp.enabledCompanies = enabled;
}

function switchTab(tabId) {
  state.currentTab = tabId;
  document.querySelectorAll('.nav-link').forEach(l => {
    l.classList.toggle('active', l.dataset.tab === tabId);
  });
  document.querySelectorAll('.content-section').forEach(sec => {
    sec.classList.toggle('active', sec.id === `section-${tabId}`);
  });
  
  if (tabId === 'price-sentiment') {
    renderPriceVsSentimentChart();
  } else if (tabId === 'aapl-amd') {
    renderPairSection();
  } else if (tabId === 'cross-company') {
    renderMultiCompanySection();
    renderCrossCompanyCharts();
  } else if (tabId === 'ratios') {
    renderRatioCharts();
  }
}

// Master Render
function renderAll() {
  updateHeroBanner();
  renderOverviewCards();
  renderFinancialStatements();
  renderFinancialRatios();
  renderSentimentSection();
  renderPriceVsSentimentSection();
  renderPairSection();
  renderMultiCompanySection();
  renderCrossCompanyCharts();
}

// 1. Hero Banner
function updateHeroBanner() {
  const isComp = state.currentCompany === 'COMP';
  const company = PROFILES[state.currentCompany] || PROFILES.AAPL;
  
  const titleEl = document.getElementById('hero-title');
  const descEl = document.getElementById('hero-desc');
  const metaContainer = document.getElementById('hero-meta');

  if (isComp) {
    const aaplMetrics = state.data.comparison.metrics.AAPL;
    const amdMetrics = state.data.comparison.metrics.AMD;
    const aaplLatest = aaplMetrics[aaplMetrics.length - 1];
    const amdLatest = amdMetrics[amdMetrics.length - 1];
    const aaplPriceChange = ((aaplLatest.split_adj_close - aaplMetrics[0].split_adj_close) / aaplMetrics[0].split_adj_close) * 100;
    const amdPriceChange = ((amdLatest.split_adj_close - amdMetrics[0].split_adj_close) / amdMetrics[0].split_adj_close) * 100;

    let compTitle = `Apple vs. AMD <span class="ticker-tag" style="background:#a855f7; color:#fff;">AAPL vs AMD</span>`;
    let compMeta = `
      <div class="meta-item">
        <div class="meta-label">Apple FY25 Revenue</div>
        <div class="meta-val" style="color:var(--apple-color);">${formatMillions(aaplLatest.net_sales_m)}</div>
      </div>
      <div class="meta-item">
        <div class="meta-label">AMD FY25 Revenue</div>
        <div class="meta-val" style="color:var(--amd-color);">${formatMillions(amdLatest.net_sales_m)}</div>
      </div>
    `;

    if (state.customCompany) {
      const customMetrics = state.customCompany.metrics;
      const customLatest = customMetrics[customMetrics.length - 1];
      const customPriceChange = ((customLatest.split_adj_close - customMetrics[0].split_adj_close) / customMetrics[0].split_adj_close) * 100;
      compTitle = `Apple vs. AMD vs. ${state.customCompany.ticker} <span class="ticker-tag" style="background:#10b981; color:#0b1120;">3-Way Analysis</span>`;
      compMeta += `
        <div class="meta-item">
          <div class="meta-label">${state.customCompany.ticker} FY25 Revenue</div>
          <div class="meta-val" style="color:var(--accent-emerald);">${formatMillions(customLatest.net_sales_m)}</div>
        </div>
        <div class="meta-item">
          <div class="meta-label">10-Yr Equity Returns</div>
          <div class="meta-val">+${aaplPriceChange.toFixed(0)}% | +${amdPriceChange.toFixed(0)}% | +${customPriceChange.toFixed(0)}%</div>
        </div>
      `;
    } else {
      compMeta += `
        <div class="meta-item">
          <div class="meta-label">10-Yr Stock Return</div>
          <div class="meta-val">+${aaplPriceChange.toFixed(0)}% vs +${amdPriceChange.toFixed(0)}%</div>
        </div>
      `;
    }

    titleEl.innerHTML = compTitle;
    descEl.textContent = 'Ten-Year Longitudinal Comparison: Financial Performance, Profitability, Asset Velocity & MD&A Sentiment (FY2016 - FY2025).';
    metaContainer.innerHTML = compMeta;
  } else {
    const metrics = (state.currentCompany === 'CUSTOM' && state.customCompany) 
      ? state.customCompany.metrics 
      : state.data.comparison.metrics[state.currentCompany];
      
    const latest = metrics[metrics.length - 1];
    const fy16 = metrics[0];
    const priceChange = ((latest.split_adj_close - fy16.split_adj_close) / fy16.split_adj_close) * 100;
    
    titleEl.innerHTML = `${company.name} <span class="ticker-tag" style="color:${company.color}">${company.ticker}</span>`;
    descEl.textContent = `${company.description} | CIK: ${company.cik} | Fiscal Year-End: ${company.fiscalEnd || 'Annual'}.`;
    
    metaContainer.innerHTML = `
      <div class="meta-item">
        <div class="meta-label">FY2025 Revenue</div>
        <div class="meta-val">${formatMillions(latest.net_sales_m)}</div>
      </div>
      <div class="meta-item">
        <div class="meta-label">FY2025 Net Income</div>
        <div class="meta-val">${formatMillions(latest.net_income_m)}</div>
      </div>
      <div class="meta-item">
        <div class="meta-label">Split-Adj Close (FY25)</div>
        <div class="meta-val">${formatCurrency(latest.split_adj_close)}</div>
      </div>
      <div class="meta-item">
        <div class="meta-label">10-Yr Stock Return</div>
        <div class="meta-val" style="color:var(--accent-emerald);">+${priceChange.toFixed(1)}%</div>
      </div>
    `;
  }
}

// 2. Overview Stat Cards
function renderOverviewCards() {
  const container = document.getElementById('overview-cards');
  const targetTicker = state.currentCompany === 'COMP' ? 'AAPL' : state.currentCompany;
  const metrics = (targetTicker === 'CUSTOM' && state.customCompany)
    ? state.customCompany.metrics
    : state.data.comparison.metrics[targetTicker];
  
  if (!metrics) return;
  const latest = metrics[metrics.length - 1];

    const hasSentiment = latest.pn_ratio !== null && latest.pn_ratio !== undefined && latest.sentiment_verified !== false;
    let sentimentBadge = `<span class="stat-badge ${latest.pn_ratio > 1.0 ? 'green' : 'amber'}">${latest.pn_ratio > 1.0 ? 'Net Positive' : 'Net Negative'}</span>`;
    let sentimentVal = latest.pn_ratio ? latest.pn_ratio.toFixed(4) : '—';
    let sentimentSub = `Pos: ${latest.pos_words || 0} | Neg: ${latest.neg_words || 0} words`;

    if (!hasSentiment) {
      sentimentBadge = `<span class="stat-badge" style="background:rgba(148,163,184,0.2); color:#94a3b8;">Unavailable</span>`;
      sentimentVal = '—';
      sentimentSub = 'Verified for AAPL & AMD only';
    }

    container.innerHTML = `
    <div class="stat-card">
      <div class="stat-header">
        <span class="stat-title">FY2025 Total Net Sales</span>
        <span class="stat-badge green">${latest.revenue_growth_pct ? formatPercent(latest.revenue_growth_pct) + ' YoY' : 'N/A'}</span>
      </div>
      <div class="stat-value">${formatMillions(latest.net_sales_m)}</div>
      <div class="stat-subtext">FY2016 Baseline: ${formatMillions(metrics[0].net_sales_m)}</div>
    </div>

    <div class="stat-card">
      <div class="stat-header">
        <span class="stat-title">FY2025 Net Income</span>
        <span class="stat-badge green">${latest.net_income_growth_pct ? formatPercent(latest.net_income_growth_pct) + ' YoY' : 'N/A'}</span>
      </div>
      <div class="stat-value">${formatMillions(latest.net_income_m)}</div>
      <div class="stat-subtext">Net Profit Margin: ${formatPercent(latest.net_profit_margin)}</div>
    </div>

    <div class="stat-card">
      <div class="stat-header">
        <span class="stat-title">Split-Adjusted Stock Price</span>
        <span class="stat-badge">${latest.indexed_price ? 'Index: ' + latest.indexed_price.toFixed(1) : 'Market'}</span>
      </div>
      <div class="stat-value">${formatCurrency(latest.split_adj_close)}</div>
      <div class="stat-subtext">Trading Date: ${latest.trading_date} (FY16: ${formatCurrency(metrics[0].split_adj_close)})</div>
    </div>

    <div class="stat-card">
      <div class="stat-header">
        <span class="stat-title">MD&A P/N Sentiment Ratio</span>
        ${sentimentBadge}
      </div>
      <div class="stat-value">${sentimentVal}</div>
      <div class="stat-subtext">${sentimentSub}</div>
    </div>
  `;
}

// 3. Financial Statements View
function renderFinancialStatements() {
  const container = document.getElementById('statements-table-container');
  const targetTicker = state.currentCompany === 'COMP' ? 'AAPL' : state.currentCompany;
  const companyRaw = (targetTicker === 'CUSTOM' && state.customCompany)
    ? state.customCompany
    : state.data[targetTicker.toLowerCase()];
  
  if (!companyRaw) return;

  let rawSheet = [];
  let title = '';
  if (state.currentStatement === 'is') {
    rawSheet = companyRaw.income_statement;
    title = `${companyRaw.ticker || targetTicker} Income Statement (FY2016 - FY2025)`;
  } else if (state.currentStatement === 'bs') {
    rawSheet = companyRaw.balance_sheet;
    title = `${companyRaw.ticker || targetTicker} Balance Sheet (FY2016 - FY2025)`;
  } else {
    rawSheet = companyRaw.cash_flows;
    title = `${companyRaw.ticker || targetTicker} Statement of Cash Flows (FY2016 - FY2025)`;
  }

  // Synthesize statements directly from metrics if pre-parsed 2D array is absent
  if ((!rawSheet || rawSheet.length === 0) && companyRaw.metrics) {
    const years = companyRaw.metrics.map(m => m.fiscal_year);
    if (state.currentStatement === 'is') {
      rawSheet = [
        ['Line Item', ...years],
        ['Total Net Sales / Revenue ($M)', ...companyRaw.metrics.map(m => m.net_sales_m)],
        ['Revenue Growth YoY (%)', ...companyRaw.metrics.map(m => m.revenue_growth_pct !== null && m.revenue_growth_pct !== undefined ? (m.revenue_growth_pct * 100).toFixed(1) + '%' : '—')],
        ['Net Income ($M)', ...companyRaw.metrics.map(m => m.net_income_m)],
        ['Net Income Growth YoY (%)', ...companyRaw.metrics.map(m => m.net_income_growth_pct !== null && m.net_income_growth_pct !== undefined ? (m.net_income_growth_pct * 100).toFixed(1) + '%' : '—')],
        ['Net Profit Margin (%)', ...companyRaw.metrics.map(m => m.net_profit_margin !== null && m.net_profit_margin !== undefined ? (m.net_profit_margin * 100).toFixed(1) + '%' : '—')]
      ];
    } else if (state.currentStatement === 'bs') {
      rawSheet = [
        ['Line Item', ...years],
        ['Current Assets ($M)', ...companyRaw.metrics.map(m => m.current_assets_m)],
        ['Total Assets ($M)', ...companyRaw.metrics.map(m => m.total_assets_m)],
        ['Current Liabilities ($M)', ...companyRaw.metrics.map(m => m.current_liabilities_m)],
        ['Total Liabilities ($M)', ...companyRaw.metrics.map(m => m.total_liabilities_m)],
        ['Stockholders\' Equity ($M)', ...companyRaw.metrics.map(m => m.stockholders_equity_m)],
        ['Total Liabilities and Stockholders\' Equity ($M)', ...companyRaw.metrics.map(m => (m.total_liabilities_m + m.stockholders_equity_m + (m.temporary_equity_m || 0)))],
        ['Balance Sheet Difference ($M)', ...companyRaw.metrics.map(m => m.balance_sheet_diff_m.toFixed(1))]
      ];
    } else {
      rawSheet = [
        ['Line Item', ...years],
        ['Cash Provided by Operating Activities ($M)', ...companyRaw.metrics.map(m => m.operating_cf_m)]
      ];
    }
  }

  document.getElementById('statement-title-text').textContent = title;

  if (!rawSheet || rawSheet.length === 0) {
    container.innerHTML = '<p class="text-muted" style="padding:1rem;">No statement data available.</p>';
    return;
  }

  let html = '<table class="data-table"><thead><tr>';
  const headerRow = rawSheet[0] && Array.isArray(rawSheet[0]) && rawSheet[0][0] === 'Line Item' ? rawSheet[0] : (rawSheet[2] || rawSheet[1]);
  
  headerRow.forEach((col, idx) => {
    html += `<th class="${idx > 0 ? 'num' : ''}">${col || ''}</th>`;
  });
  html += '</tr></thead><tbody>';

  const startRow = rawSheet[0] && Array.isArray(rawSheet[0]) && rawSheet[0][0] === 'Line Item' ? 1 : 3;

  for (let r = startRow; r < rawSheet.length; r++) {
    const row = rawSheet[r];
    const isHeading = row.slice(1).every(cell => cell === null || cell === undefined || cell === '' || cell === '—');
    const isTotal = String(row[0]).toLowerCase().includes('total') || String(row[0]).toLowerCase().includes('net income');
    
    html += `<tr class="${isHeading ? 'heading-row' : ''} ${isTotal ? 'total-row' : ''}">`;
    row.forEach((cell, idx) => {
      if (idx === 0) {
        html += `<td class="${isTotal ? 'bold' : ''}" style="${isHeading ? 'font-weight:700; color:var(--accent-blue);' : ''}">${cell || ''}</td>`;
      } else {
        const numVal = typeof cell === 'number' ? cell.toLocaleString('en-US') : (cell || '—');
        html += `<td class="num ${isTotal ? 'bold' : ''}">${numVal}</td>`;
      }
    });
    html += '</tr>';
  }
  html += '</tbody></table>';
  container.innerHTML = html;
}

// 4. Financial Ratios View
function renderFinancialRatios() {
  const targetTicker = state.currentCompany === 'COMP' ? 'AAPL' : state.currentCompany;
  const ratios = (targetTicker === 'CUSTOM' && state.customCompany)
    ? state.customCompany.ratios
    : state.data[targetTicker.toLowerCase()].ratios;

  const cardsContainer = document.getElementById('ratio-cards-container');
  const tableContainer = document.getElementById('ratios-table-container');

  if (!ratios) return;

  const crLast = ratios.current_ratio.values[ratios.current_ratio.values.length - 1];
  const ltaLast = ratios.liabilities_to_assets.values[ratios.liabilities_to_assets.values.length - 1];
  const atoLast = ratios.total_asset_turnover.values[ratios.total_asset_turnover.values.length - 1];
  const npmLast = ratios.net_profit_margin.values[ratios.net_profit_margin.values.length - 1];

  cardsContainer.innerHTML = `
    <div class="ratio-card">
      <div class="ratio-name">
        <span>${ratios.current_ratio.name}</span>
        <span class="ratio-category stat-badge green">${ratios.current_ratio.category}</span>
      </div>
      <div class="ratio-formula">${ratios.current_ratio.formula}</div>
      <div class="ratio-benchmark">Benchmark: ${ratios.current_ratio.benchmark}</div>
      <div class="ratio-latest">
        <div>
          <div class="ratio-val">${crLast ? crLast.toFixed(2) + 'x' : '—'}</div>
          <div class="ratio-avg">10-Yr Avg: ${ratios.current_ratio.ten_year_avg ? ratios.current_ratio.ten_year_avg.toFixed(2) + 'x' : '—'}</div>
        </div>
        <div style="font-size:0.8rem; color:${crLast >= 1.0 ? 'var(--accent-emerald)' : 'var(--accent-amber)'}">
          ${crLast >= 1.0 ? 'Working capital buffer' : 'Aggressive working capital'}
        </div>
      </div>
    </div>

    <div class="ratio-card">
      <div class="ratio-name">
        <span>${ratios.liabilities_to_assets.name}</span>
        <span class="ratio-category stat-badge amber">${ratios.liabilities_to_assets.category}</span>
      </div>
      <div class="ratio-formula">${ratios.liabilities_to_assets.formula}</div>
      <div class="ratio-benchmark">Benchmark: ${ratios.liabilities_to_assets.benchmark}</div>
      <div class="ratio-latest">
        <div>
          <div class="ratio-val">${ltaLast ? (ltaLast * 100).toFixed(1) + '%' : '—'}</div>
          <div class="ratio-avg">10-Yr Avg: ${ratios.liabilities_to_assets.ten_year_avg ? (ratios.liabilities_to_assets.ten_year_avg * 100).toFixed(1) + '%' : '—'}</div>
        </div>
        <div style="font-size:0.8rem; color:var(--text-dim);">Capital Structure</div>
      </div>
    </div>

    <div class="ratio-card">
      <div class="ratio-name">
        <span>${ratios.total_asset_turnover.name}</span>
        <span class="ratio-category stat-badge">${ratios.total_asset_turnover.category}</span>
      </div>
      <div class="ratio-formula">${ratios.total_asset_turnover.formula}</div>
      <div class="ratio-benchmark">Benchmark: ${ratios.total_asset_turnover.benchmark}</div>
      <div class="ratio-latest">
        <div>
          <div class="ratio-val">${atoLast ? atoLast.toFixed(2) + 'x' : '—'}</div>
          <div class="ratio-avg">10-Yr Avg: ${ratios.total_asset_turnover.ten_year_avg ? ratios.total_asset_turnover.ten_year_avg.toFixed(2) + 'x' : '—'}</div>
        </div>
        <div style="font-size:0.8rem; color:var(--accent-emerald);">Asset Productivity</div>
      </div>
    </div>

    <div class="ratio-card">
      <div class="ratio-name">
        <span>${ratios.net_profit_margin.name}</span>
        <span class="ratio-category stat-badge green">${ratios.net_profit_margin.category}</span>
      </div>
      <div class="ratio-formula">${ratios.net_profit_margin.formula}</div>
      <div class="ratio-benchmark">Benchmark: ${ratios.net_profit_margin.benchmark}</div>
      <div class="ratio-latest">
        <div>
          <div class="ratio-val">${npmLast ? (npmLast * 100).toFixed(1) + '%' : '—'}</div>
          <div class="ratio-avg">10-Yr Avg: ${ratios.net_profit_margin.ten_year_avg ? (ratios.net_profit_margin.ten_year_avg * 100).toFixed(1) + '%' : '—'}</div>
        </div>
        <div style="font-size:0.8rem; color:var(--accent-emerald);">Bottom-Line Return</div>
      </div>
    </div>
  `;

  // Render 10-Year Ratio Table
  let html = `<table class="data-table"><thead><tr>
    <th>Financial Ratio</th>
    <th>Standard Definition</th>`;
  state.data.comparison.fiscal_years.forEach(yr => { html += `<th class="num">${yr}</th>`; });
  html += `<th class="num">10-Yr Avg</th></tr></thead><tbody>`;

  const ratioList = [ratios.current_ratio, ratios.liabilities_to_assets, ratios.total_asset_turnover, ratios.net_profit_margin];
  ratioList.forEach(r => {
    html += `<tr>
      <td class="bold">${r.name}</td>
      <td style="color:var(--text-muted); font-size:0.8rem;">${r.formula}</td>`;
    r.values.forEach(v => {
      const formatted = v !== null && v !== undefined
        ? (r.name.includes('Margin') || r.name.includes('Assets') ? (v * 100).toFixed(1) + '%' : v.toFixed(2) + 'x')
        : '—';
      html += `<td class="num">${formatted}</td>`;
    });
    const avgFormatted = r.ten_year_avg !== null && r.ten_year_avg !== undefined
      ? (r.name.includes('Margin') || r.name.includes('Assets') ? (r.ten_year_avg * 100).toFixed(1) + '%' : r.ten_year_avg.toFixed(2) + 'x')
      : '—';
    html += `<td class="num bold" style="color:var(--accent-blue);">${avgFormatted}</td></tr>`;
  });
  html += '</tbody></table>';
  tableContainer.innerHTML = html;

  renderRatioCharts();
}

function renderRatioCharts() {
  const canvas = document.getElementById('chart-ratios');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const targetTicker = state.currentCompany === 'COMP' ? 'AAPL' : state.currentCompany;
  const ratios = (targetTicker === 'CUSTOM' && state.customCompany)
    ? state.customCompany.ratios
    : state.data[targetTicker.toLowerCase()].ratios;

  if (state.ratioChartInstance) {
    state.ratioChartInstance.destroy();
  }

  state.ratioChartInstance = new Chart(ctx, {
    type: 'line',
    data: {
      labels: state.data.comparison.fiscal_years,
      datasets: [
        {
          label: 'Current Ratio (x)',
          data: ratios.current_ratio.values,
          borderColor: '#10b981',
          backgroundColor: '#10b981',
          borderWidth: 2.5,
          pointRadius: 4,
          tension: 0.1
        },
        {
          label: 'Total Asset Turnover (x)',
          data: ratios.total_asset_turnover.values,
          borderColor: '#38bdf8',
          backgroundColor: '#38bdf8',
          borderWidth: 2.5,
          pointRadius: 4,
          tension: 0.1
        },
        {
          label: 'Net Profit Margin (%)',
          data: ratios.net_profit_margin.values.map(v => v !== null ? v * 100 : null),
          borderColor: '#f59e0b',
          backgroundColor: '#f59e0b',
          borderWidth: 2.5,
          pointRadius: 4,
          tension: 0.1
        },
        {
          label: 'Liabilities-to-Assets (%)',
          data: ratios.liabilities_to_assets.values.map(v => v !== null ? v * 100 : null),
          borderColor: '#ef4444',
          backgroundColor: '#ef4444',
          borderWidth: 2.5,
          pointRadius: 4,
          tension: 0.1
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: { color: '#94a3b8', font: { size: 12 } }
        },
        tooltip: {
          backgroundColor: '#1e293b',
          titleColor: '#fff',
          bodyColor: '#cbd5e1',
          borderColor: '#334155',
          borderWidth: 1
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(255, 255, 255, 0.05)' },
          ticks: { color: '#94a3b8' }
        },
        y: {
          grid: { color: 'rgba(255, 255, 255, 0.05)' },
          ticks: { color: '#94a3b8' }
        }
      }
    }
  });
}

// 5. MD&A Sentiment Section
function renderSentimentSection() {
  const targetTicker = state.currentCompany === 'COMP' ? 'AAPL' : state.currentCompany;
  const metrics = (targetTicker === 'CUSTOM' && state.customCompany)
    ? state.customCompany.metrics
    : state.data.comparison.metrics[targetTicker];

  const matchedWords = (targetTicker === 'CUSTOM' && state.customCompany)
    ? state.customCompany.matched_words
    : state.data[targetTicker.toLowerCase()].matched_words;

  const tableContainer = document.getElementById('sentiment-table-container');
  const wordsContainer = document.getElementById('matched-words-container');

  let html = `<table class="data-table"><thead><tr>
    <th>Fiscal Year</th>
    <th class="num">MD&A Word Count</th>
    <th class="num">Positive Words</th>
    <th class="num">Negative Words</th>
    <th class="num">Pos / 1k Words</th>
    <th class="num">Neg / 1k Words</th>
    <th class="num">P/N Ratio</th>
    <th class="num">Net Sent / 1k</th>
  </tr></thead><tbody>`;

  metrics.forEach(m => {
    const isNetPos = m.net_sentiment_k > 0;
    html += `<tr>
      <td class="bold">${m.fiscal_year}</td>
      <td class="num">${formatNumber(m.total_word_count)}</td>
      <td class="num" style="color:var(--accent-emerald);">${m.pos_words !== null && m.pos_words !== undefined ? m.pos_words : '—'}</td>
      <td class="num" style="color:var(--accent-crimson);">${m.neg_words !== null && m.neg_words !== undefined ? m.neg_words : '—'}</td>
      <td class="num">${m.pos_rate_k ? m.pos_rate_k.toFixed(2) : '—'}</td>
      <td class="num">${m.neg_rate_k ? m.neg_rate_k.toFixed(2) : '—'}</td>
      <td class="num bold" style="color:${m.pn_ratio > 1.0 ? 'var(--accent-emerald)' : 'var(--accent-amber)'};">${m.pn_ratio ? m.pn_ratio.toFixed(4) : '—'}</td>
      <td class="num bold" style="color:${isNetPos ? 'var(--accent-emerald)' : 'var(--accent-crimson)'};">${m.net_sentiment_k !== null ? ((m.net_sentiment_k >= 0 ? '+' : '') + m.net_sentiment_k.toFixed(2)) : '—'}</td>
    </tr>`;
  });
  html += '</tbody></table>';
  tableContainer.innerHTML = html;

  // Outlier review card
  const outlierCard = document.getElementById('outlier-audit-card');
  if (targetTicker === 'AAPL') {
    outlierCard.innerHTML = `
      <div class="callout info">
        <h4>Audit Review: Apple FY2021 Sentiment Spike (P/N Ratio: 0.9286x)</h4>
        <p>In FY2021, Apple recorded a historic jump in narrative optimism: <strong>13 positive words vs. only 14 negative words</strong> (P/N ratio of 0.9286x vs. 10-year median of 0.23x). This surge corresponded directly with the peak pandemic hardware upgrade cycle, which drove Apple's annual revenue up +33.3% YoY to $365,817M. However, despite this optimism, management maintained strict SEC legal conservativism with standardized cautionary risk boilerplate.</p>
      </div>
    `;
  } else if (targetTicker === 'AMD') {
    outlierCard.innerHTML = `
      <div class="callout info">
        <h4>Audit Review: AMD FY2020 Operational Turnaround Spike (P/N Ratio: 1.3611x)</h4>
        <p>In FY2020, AMD achieved a dramatic multi-year turnaround: <strong>49 positive words vs. 36 negative words</strong> (P/N ratio 1.3611x, Net Sentiment +2.54 per 1k words). Narrative enthusiasm reflected the successful execution of the Zen 3 / EPYC Milan microarchitecture roadmap and gross margin expansion, coinciding with AMD stock skyrocketing from $11.34 (FY16) to $91.81 (FY20).</p>
      </div>
    `;
  } else {
    outlierCard.innerHTML = `
      <div class="callout info" style="border-color:var(--accent-emerald);">
        <h4 style="color:var(--accent-emerald);">MD&A Sentiment Status: ${state.customCompany ? state.customCompany.company_name : 'Microsoft Corp.'} (${state.customCompany ? state.customCompany.ticker : 'MSFT'})</h4>
        <p>Loughran-McDonald MD&A sentiment metrics are verified for primary research subjects <strong>Apple Inc. (AAPL)</strong> and <strong>AMD (AMD)</strong>. Microsoft Corporation (MSFT) is included across this platform as an audited 10-year US-GAAP financial statement and ratio peer benchmark. Longitudinal MD&A sentiment metrics for MSFT are designated as unavailable (<code>—</code>) in accordance with accounting conservatism standards.</p>
      </div>
    `;
  }

  // Render Top Matched Words Table
  if (matchedWords && matchedWords.length > 1) {
    let wHtml = `<table class="data-table"><thead><tr>
      <th>Matched Word</th>
      <th>Sentiment Category</th>
      <th class="num">Frequency Count</th>
    </tr></thead><tbody>`;
    for (let r = 1; r < Math.min(26, matchedWords.length); r++) {
      const row = matchedWords[r];
      const isPos = String(row[1]).toLowerCase().includes('positive');
      wHtml += `<tr>
        <td class="bold font-mono">${row[0]}</td>
        <td><span class="stat-badge ${isPos ? 'green' : 'red'}">${row[1]}</span></td>
        <td class="num">${row[2]}</td>
      </tr>`;
    }
    wHtml += '</tbody></table>';
    wordsContainer.innerHTML = wHtml;
  } else {
    wordsContainer.innerHTML = `
      <div style="padding:1.5rem; text-align:center; color:var(--text-muted); font-style:italic;">
        Top matched Loughran-McDonald dictionary words unavailable for Microsoft (MSFT). MD&A sentiment analysis is verified for primary research subjects Apple Inc. (AAPL) and AMD.
      </div>
    `;
  }
}

// 6. Stock Price vs MD&A Sentiment Section
function renderPriceVsSentimentSection() {
  const targetTicker = state.currentCompany === 'COMP' ? 'AAPL' : state.currentCompany;
  const metrics = (targetTicker === 'CUSTOM' && state.customCompany)
    ? state.customCompany.metrics
    : state.data.comparison.metrics[targetTicker];

  const tableContainer = document.getElementById('price-sentiment-table-container');

  let html = `<table class="data-table"><thead><tr>
    <th>Fiscal Year</th>
    <th>Fiscal Year-End</th>
    <th>Market Trading Date</th>
    <th class="num">Split-Adjusted Close</th>
    <th class="num">Indexed Price (FY16=100)</th>
    <th class="num">Positive Words</th>
    <th class="num">Negative Words</th>
    <th class="num">P/N Sentiment Ratio</th>
    <th class="num">Net Sent / 1k</th>
    <th>Market Timing Note</th>
  </tr></thead><tbody>`;

  metrics.forEach(m => {
    html += `<tr>
      <td class="bold">${m.fiscal_year}</td>
      <td style="color:var(--text-muted);">${m.fiscal_end_date}</td>
      <td style="color:var(--text-muted);">${m.trading_date}</td>
      <td class="num bold">${formatCurrency(m.split_adj_close)}</td>
      <td class="num">${m.indexed_price !== null ? m.indexed_price.toFixed(1) : '—'}</td>
      <td class="num" style="color:var(--accent-emerald);">${m.pos_words !== null && m.pos_words !== undefined ? m.pos_words : '—'}</td>
      <td class="num" style="color:var(--accent-crimson);">${m.neg_words !== null && m.neg_words !== undefined ? m.neg_words : '—'}</td>
      <td class="num bold" style="color:${m.pn_ratio > 1.0 ? 'var(--accent-emerald)' : 'var(--accent-amber)'};">${m.pn_ratio ? m.pn_ratio.toFixed(4) : '—'}</td>
      <td class="num">${m.net_sentiment_k !== null ? ((m.net_sentiment_k >= 0 ? '+' : '') + m.net_sentiment_k.toFixed(2)) : '—'}</td>
      <td style="font-size:0.8rem; color:var(--text-dim);">${m.timing_note}</td>
    </tr>`;
  });
  html += '</tbody></table>';
  tableContainer.innerHTML = html;

  renderPriceVsSentimentChart();
}

function renderPriceVsSentimentChart() {
  const canvas = document.getElementById('chart-price-sentiment');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const targetTicker = state.currentCompany === 'COMP' ? 'AAPL' : state.currentCompany;
  const metrics = (targetTicker === 'CUSTOM' && state.customCompany)
    ? state.customCompany.metrics
    : state.data.comparison.metrics[targetTicker];

  if (state.chartInstance) {
    state.chartInstance.destroy();
  }

  const hasSentiment = metrics && metrics.some(m => m.pn_ratio !== null && m.pn_ratio !== undefined);

  const chartDatasets = [
    {
      label: 'Split-Adjusted Stock Price ($)',
      data: metrics.map(m => m.split_adj_close),
      borderColor: '#1d4ed8',
      backgroundColor: '#1d4ed8',
      borderWidth: 2.5,
      pointRadius: 5,
      pointHoverRadius: 7,
      pointStyle: 'circle',
      yAxisID: 'yLeft',
      tension: 0
    }
  ];

  if (hasSentiment) {
    chartDatasets.push({
      label: 'MD&A Positive / Negative Ratio',
      data: metrics.map(m => m.pn_ratio),
      borderColor: '#ef4444',
      backgroundColor: '#ef4444',
      borderWidth: 2.5,
      pointRadius: 5,
      pointHoverRadius: 7,
      pointStyle: 'rect',
      yAxisID: 'yRight',
      tension: 0
    });
  }

  state.chartInstance = new Chart(ctx, {
    type: 'line',
    data: {
      labels: metrics.map(m => m.fiscal_year),
      datasets: chartDatasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: { color: '#94a3b8', font: { size: 12 }, padding: 20 }
        },
        tooltip: {
          backgroundColor: '#1e293b',
          titleColor: '#fff',
          bodyColor: '#cbd5e1',
          borderColor: '#334155',
          borderWidth: 1,
          callbacks: {
            label: function(context) {
              const label = context.dataset.label || '';
              const val = context.parsed.y;
              if (val === null || val === undefined || isNaN(val)) {
                return `${label}: Unavailable`;
              }
              if (context.dataset.yAxisID === 'yLeft') {
                return `${label}: $${val.toFixed(2)}`;
              } else {
                return `${label}: ${val.toFixed(4)}x`;
              }
            }
          }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { color: '#94a3b8' }
        },
        yLeft: {
          type: 'linear',
          position: 'left',
          title: {
            display: true,
            text: 'Split-Adjusted Stock Price ($)',
            color: '#94a3b8',
            font: { weight: 'bold' }
          },
          grid: { color: 'rgba(255, 255, 255, 0.08)' },
          ticks: {
            color: '#94a3b8',
            callback: (v) => '$' + v
          }
        },
        yRight: {
          display: hasSentiment,
          type: 'linear',
          position: 'right',
          title: {
            display: hasSentiment,
            text: 'MD&A Positive / Negative Ratio',
            color: '#94a3b8',
            font: { weight: 'bold' }
          },
          grid: { display: false },
          ticks: {
            color: '#94a3b8',
            callback: (v) => v.toFixed(2)
          }
        }
      }
    }
  });
}

// 6B. Dedicated Apple vs. AMD Pair Comparison Section
function renderPairSection() {
  const metricKey = state.pairComp.selectedMetric;
  const metricInfo = METRIC_DEFINITIONS[metricKey] || METRIC_DEFINITIONS.net_sales_m;

  // Update Metric Description Card
  const descCard = document.getElementById('metric-pair-desc-card');
  if (descCard) {
    descCard.innerHTML = `
      <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:0.5rem;">
        <div>
          <strong style="color:#fff; font-size:0.95rem;">${metricInfo.name}</strong> 
          <span class="stat-badge" style="margin-left:0.5rem;">${metricInfo.category}</span>
        </div>
        ${metricInfo.benchmark ? `<span style="color:var(--accent-emerald); font-size:0.8rem; font-weight:600;">${metricInfo.benchmark}</span>` : ''}
      </div>
      <p style="margin-top:0.4rem; color:var(--text-muted); font-size:0.85rem;">${metricInfo.definition}</p>
      <p style="margin-top:0.25rem; color:var(--text-dim); font-size:0.8rem; font-style:italic;"><strong>Accounting Context:</strong> ${metricInfo.interpret}</p>
    `;
  }

  // Filter Fiscal Years by Range
  const allYears = state.data.comparison.fiscal_years;
  const startIdx = allYears.indexOf(state.pairComp.startYear);
  const endIdx = allYears.indexOf(state.pairComp.endYear);
  const filteredYears = allYears.slice(Math.max(0, startIdx), endIdx >= startIdx ? endIdx + 1 : allYears.length);

  // Datasets for Apple and AMD ONLY
  const compConfigs = [
    { key: 'AAPL', name: 'Apple Inc. (AAPL)', color: '#38bdf8', metrics: state.data.comparison.metrics.AAPL },
    { key: 'AMD', name: 'Advanced Micro Devices (AMD)', color: '#ef4444', metrics: state.data.comparison.metrics.AMD }
  ];

  const datasets = compConfigs.map(cfg => {
    const dataPoints = filteredYears.map(yr => {
      const rec = cfg.metrics.find(m => m.fiscal_year === yr);
      if (!rec) return null;
      let val = rec[metricKey];
      if (metricKey.includes('pct') || metricKey.includes('margin') || metricKey.includes('to_assets')) {
        return val !== null && val !== undefined ? val * 100.0 : null;
      }
      return val !== null && val !== undefined ? val : null;
    });

    return {
      label: cfg.name,
      data: dataPoints,
      borderColor: cfg.color,
      backgroundColor: state.pairComp.chartType === 'bar' ? cfg.color + 'B3' : cfg.color,
      borderWidth: 2.5,
      pointRadius: 4,
      pointHoverRadius: 6,
      tension: 0.1
    };
  });

  // Render Primary Pair Chart
  const canvas = document.getElementById('chart-pair-comparison');
  if (canvas) {
    const ctx = canvas.getContext('2d');
    if (state.pairComp.chartInstance) {
      state.pairComp.chartInstance.destroy();
    }

    state.pairComp.chartInstance = new Chart(ctx, {
      type: state.pairComp.chartType,
      data: {
        labels: filteredYears,
        datasets: datasets
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: { color: '#94a3b8', font: { size: 12 }, padding: 15 }
          },
          tooltip: {
            backgroundColor: '#1e293b',
            titleColor: '#fff',
            bodyColor: '#cbd5e1',
            borderColor: '#334155',
            borderWidth: 1,
            callbacks: {
              label: (context) => {
                const label = context.dataset.label || '';
                const val = context.parsed.y;
                if (val === null || val === undefined || isNaN(val)) {
                  return `${label}: —`;
                }
                let formatted = val;
                if (metricKey.includes('pct') || metricKey.includes('margin') || metricKey.includes('to_assets')) {
                  formatted = (val >= 0 ? '+' : '') + val.toFixed(1) + '%';
                } else if (metricKey.includes('sales') || metricKey.includes('income') || metricKey.includes('assets') || metricKey.includes('cf')) {
                  formatted = formatMillions(val);
                } else if (metricKey.includes('close')) {
                  formatted = formatCurrency(val);
                } else if (metricKey.includes('ratio') || metricKey.includes('turnover')) {
                  formatted = val.toFixed(4) + 'x';
                }
                return `${label}: ${formatted}`;
              }
            }
          }
        },
        scales: {
          x: {
            grid: { color: 'rgba(255, 255, 255, 0.05)' },
            ticks: { color: '#94a3b8' }
          },
          y: {
            grid: { color: 'rgba(255, 255, 255, 0.08)' },
            ticks: {
              color: '#94a3b8',
              callback: (v) => {
                if (metricKey.includes('pct') || metricKey.includes('margin') || metricKey.includes('to_assets')) return v.toFixed(0) + '%';
                if (metricKey.includes('close')) return '$' + v;
                if (metricKey.includes('sales') || metricKey.includes('income')) return '$' + v.toLocaleString() + 'M';
                return v;
              }
            }
          }
        }
      }
    });
  }

  // Render Detailed Comparison Table
  renderPairTable(filteredYears, compConfigs, metricKey, metricInfo);

  // Render Trajectory Chart
  renderPairTrajectoryChart();
}

// Apple vs AMD Comparison Table
function renderPairTable(years, compConfigs, metricKey, metricInfo) {
  const container = document.getElementById('pair-table-container');
  if (!container) return;

  let html = `<table class="data-table"><thead><tr>
    <th>Fiscal Year</th>`;
  
  compConfigs.forEach(c => {
    html += `<th class="num" style="color:${c.color}">${c.name}</th>`;
  });
  html += '</tr></thead><tbody>';

  years.forEach(yr => {
    html += `<tr><td class="bold">${yr}</td>`;
    compConfigs.forEach(c => {
      const rec = c.metrics.find(m => m.fiscal_year === yr);
      const val = rec ? rec[metricKey] : null;
      let displayVal = '—';
      if (val !== null && val !== undefined) {
        displayVal = metricInfo.format(val);
      }
      html += `<td class="num">${displayVal}</td>`;
    });
    html += '</tr>';
  });

  // Calculate Average Row
  html += `<tr class="total-row"><td class="bold">10-Yr Average / Total</td>`;
  compConfigs.forEach(c => {
    const validVals = years
      .map(yr => {
        const rec = c.metrics.find(m => m.fiscal_year === yr);
        return rec ? rec[metricKey] : null;
      })
      .filter(v => v !== null && v !== undefined && !isNaN(v));

    let avgDisplay = '—';
    if (validVals.length > 0) {
      const avg = validVals.reduce((a, b) => a + b, 0) / validVals.length;
      avgDisplay = metricInfo.format(avg);
    }
    html += `<td class="num bold" style="color:${c.color}">${avgDisplay}</td>`;
  });
  html += '</tr></tbody></table>';

  container.innerHTML = html;
}

// Normalized Price vs Sentiment Trajectory Chart for Apple & AMD
function renderPairTrajectoryChart() {
  const canvas = document.getElementById('chart-pair-trajectory');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const aaplMetrics = state.data.comparison.metrics.AAPL;
  const amdMetrics = state.data.comparison.metrics.AMD;

  if (state.pairComp.trajectoryChartInstance) {
    state.pairComp.trajectoryChartInstance.destroy();
  }

  const datasets = [
    {
      label: 'Apple Indexed Stock Price (FY16=100)',
      data: aaplMetrics.map(m => m.indexed_price),
      borderColor: '#38bdf8',
      backgroundColor: '#38bdf8',
      borderWidth: 2.5,
      pointRadius: 4,
      tension: 0.1
    },
    {
      label: 'AMD Indexed Stock Price (FY16=100)',
      data: amdMetrics.map(m => m.indexed_price),
      borderColor: '#ef4444',
      backgroundColor: '#ef4444',
      borderWidth: 2.5,
      pointRadius: 4,
      tension: 0.1
    },
    {
      label: 'Apple MD&A P/N Ratio (x100 for scale)',
      data: aaplMetrics.map(m => m.pn_ratio * 100),
      borderColor: '#60a5fa',
      borderDash: [5, 5],
      pointRadius: 3,
      tension: 0.1
    },
    {
      label: 'AMD MD&A P/N Ratio (x100 for scale)',
      data: amdMetrics.map(m => m.pn_ratio * 100),
      borderColor: '#f87171',
      borderDash: [5, 5],
      pointRadius: 3,
      tension: 0.1
    }
  ];

  state.pairComp.trajectoryChartInstance = new Chart(ctx, {
    type: 'line',
    data: {
      labels: state.data.comparison.fiscal_years,
      datasets: datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: { color: '#94a3b8', font: { size: 12 }, padding: 15 }
        },
        tooltip: {
          backgroundColor: '#1e293b',
          titleColor: '#fff',
          bodyColor: '#cbd5e1',
          borderColor: '#334155',
          borderWidth: 1
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(255, 255, 255, 0.05)' },
          ticks: { color: '#94a3b8' }
        },
        y: {
          grid: { color: 'rgba(255, 255, 255, 0.08)' },
          ticks: { color: '#94a3b8' }
        }
      }
    }
  });
}

// Export Functions for Apple vs AMD
function exportPairPNG() {
  if (!state.pairComp.chartInstance) return;
  const imageURI = state.pairComp.chartInstance.toBase64Image();
  const link = document.createElement('a');
  link.download = `SEC_Apple_vs_AMD_${state.pairComp.selectedMetric}_${new Date().toISOString().slice(0, 10)}.png`;
  link.href = imageURI;
  link.click();
}

function exportPairCSV() {
  const metricKey = state.pairComp.selectedMetric;
  const metricInfo = METRIC_DEFINITIONS[metricKey] || METRIC_DEFINITIONS.net_sales_m;
  const allYears = state.data.comparison.fiscal_years;
  const startIdx = allYears.indexOf(state.pairComp.startYear);
  const endIdx = allYears.indexOf(state.pairComp.endYear);
  const years = allYears.slice(Math.max(0, startIdx), endIdx >= startIdx ? endIdx + 1 : allYears.length);

  const compConfigs = [
    { name: 'Apple Inc.', ticker: 'AAPL', metrics: state.data.comparison.metrics.AAPL },
    { name: 'Advanced Micro Devices', ticker: 'AMD', metrics: state.data.comparison.metrics.AMD }
  ];

  let csv = `"Fiscal Year","Apple Inc. (AAPL) - ${metricInfo.name}","Advanced Micro Devices (AMD) - ${metricInfo.name}"\r\n`;

  years.forEach(yr => {
    const aaplRec = compConfigs[0].metrics.find(m => m.fiscal_year === yr);
    const amdRec = compConfigs[1].metrics.find(m => m.fiscal_year === yr);
    const aaplVal = aaplRec && aaplRec[metricKey] !== null && aaplRec[metricKey] !== undefined ? aaplRec[metricKey] : '';
    const amdVal = amdRec && amdRec[metricKey] !== null && amdRec[metricKey] !== undefined ? amdRec[metricKey] : '';
    csv += `"${yr}",${aaplVal},${amdVal}\r\n`;
  });

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.download = `SEC_Apple_vs_AMD_${metricKey}.csv`;
  link.href = url;
  link.click();
  URL.revokeObjectURL(url);
}

// 7. Multi-Company Comparison Section (Phase 2 Enhanced)
function renderMultiCompanySection() {
  const metricKey = state.multiComp.selectedMetric;
  const metricInfo = METRIC_DEFINITIONS[metricKey] || METRIC_DEFINITIONS.net_sales_m;

  // Update Section Title & Subtitle
  const titleEl = document.getElementById('multi-comp-title');
  const subEl = document.getElementById('multi-comp-subtitle');
  if (state.customCompany) {
    if (titleEl) titleEl.textContent = `Three-Company Comparative Analytics: Apple vs. AMD vs. ${state.customCompany.ticker}`;
    if (subEl) subEl.textContent = `Side-by-side longitudinal analysis comparing Apple, AMD, and ${state.customCompany.company_name} (${state.customCompany.ticker}) across FY2016–FY2025.`;
  } else {
    if (titleEl) titleEl.textContent = `Cross-Company Comparative Analytics: Apple (AAPL) vs. AMD (AMD)`;
    if (subEl) subEl.textContent = `Side-by-side longitudinal analysis comparing Apple (AAPL) and AMD (AMD). Click 'Analyze a Third Company' to add another company.`;
  }

  // Update Metric Description Card
  const descCard = document.getElementById('metric-desc-card');
  if (descCard) {
    descCard.innerHTML = `
      <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:0.5rem;">
        <div>
          <strong style="color:#fff; font-size:0.95rem;">${metricInfo.name}</strong> 
          <span class="stat-badge" style="margin-left:0.5rem;">${metricInfo.category}</span>
        </div>
        ${metricInfo.benchmark ? `<span style="color:var(--accent-emerald); font-size:0.8rem; font-weight:600;">${metricInfo.benchmark}</span>` : ''}
      </div>
      <p style="margin-top:0.4rem; color:var(--text-muted); font-size:0.85rem;">${metricInfo.definition}</p>
      <p style="margin-top:0.25rem; color:var(--text-dim); font-size:0.8rem; font-style:italic;"><strong>Accounting Context:</strong> ${metricInfo.interpret}</p>
    `;
  }

  // Filter Fiscal Years by Range
  const allYears = state.data.comparison.fiscal_years;
  const startIdx = allYears.indexOf(state.multiComp.startYear);
  const endIdx = allYears.indexOf(state.multiComp.endYear);
  const filteredYears = allYears.slice(Math.max(0, startIdx), endIdx >= startIdx ? endIdx + 1 : allYears.length);

  // Build Chart Datasets for Active Companies
  const datasets = [];
  const compConfigs = [
    { key: 'AAPL', name: 'Apple Inc. (AAPL)', color: '#38bdf8', metrics: state.data.comparison.metrics.AAPL },
    { key: 'AMD', name: 'Advanced Micro Devices (AMD)', color: '#ef4444', metrics: state.data.comparison.metrics.AMD }
  ];

  if (state.customCompany) {
    compConfigs.push({
      key: 'CUSTOM',
      name: `${state.customCompany.company_name} (${state.customCompany.ticker})`,
      color: '#10b981',
      metrics: state.customCompany.metrics
    });
  }

  const enabled = state.multiComp.enabledCompanies;

  compConfigs.forEach(cfg => {
    if (!enabled.includes(cfg.key)) return;

    const dataPoints = filteredYears.map(yr => {
      const rec = cfg.metrics.find(m => m.fiscal_year === yr);
      if (!rec) return null;
      let val = rec[metricKey];
      if (metricKey.includes('pct') || metricKey.includes('margin') || metricKey.includes('to_assets')) {
        return val !== null && val !== undefined ? val * 100.0 : null;
      }
      return val !== null && val !== undefined ? val : null;
    });

    const hasData = dataPoints.some(v => v !== null && v !== undefined);
    if (!hasData && cfg.key === 'CUSTOM') {
      // Don't render phantom dataset for MSFT on unverified sentiment metrics
      return;
    }

    datasets.push({
      label: cfg.name,
      data: dataPoints,
      borderColor: cfg.color,
      backgroundColor: state.multiComp.chartType === 'bar' ? cfg.color + 'B3' : cfg.color,
      borderWidth: 2.5,
      pointRadius: 4,
      pointHoverRadius: 6,
      tension: 0.1
    });
  });

  // Render Chart
  const canvas = document.getElementById('chart-multi-comparison');
  if (canvas) {
    const ctx = canvas.getContext('2d');
    if (state.multiCompChartInstance) {
      state.multiCompChartInstance.destroy();
    }

    state.multiCompChartInstance = new Chart(ctx, {
      type: state.multiComp.chartType,
      data: {
        labels: filteredYears,
        datasets: datasets
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: { color: '#94a3b8', font: { size: 12 }, padding: 15 }
          },
          tooltip: {
            backgroundColor: '#1e293b',
            titleColor: '#fff',
            bodyColor: '#cbd5e1',
            borderColor: '#334155',
            borderWidth: 1,
            callbacks: {
              label: (context) => {
                const label = context.dataset.label || '';
                const val = context.parsed.y;
                if (val === null || val === undefined || isNaN(val)) {
                  return `${label}: Unavailable`;
                }
                let formatted = val;
                if (metricKey.includes('pct') || metricKey.includes('margin') || metricKey.includes('to_assets')) {
                  formatted = (val >= 0 ? '+' : '') + val.toFixed(1) + '%';
                } else if (metricKey.includes('sales') || metricKey.includes('income') || metricKey.includes('assets') || metricKey.includes('cf')) {
                  formatted = formatMillions(val);
                } else if (metricKey.includes('close')) {
                  formatted = formatCurrency(val);
                } else if (metricKey.includes('ratio') || metricKey.includes('turnover')) {
                  formatted = val.toFixed(4) + 'x';
                }
                return `${label}: ${formatted}`;
              }
            }
          }
        },
        scales: {
          x: {
            grid: { color: 'rgba(255, 255, 255, 0.05)' },
            ticks: { color: '#94a3b8' }
          },
          y: {
            grid: { color: 'rgba(255, 255, 255, 0.08)' },
            ticks: {
              color: '#94a3b8',
              callback: (v) => {
                if (metricKey.includes('pct') || metricKey.includes('margin') || metricKey.includes('to_assets')) return v.toFixed(0) + '%';
                if (metricKey.includes('close')) return '$' + v;
                if (metricKey.includes('sales') || metricKey.includes('income')) return '$' + v.toLocaleString() + 'M';
                return v;
              }
            }
          }
        }
      }
    });
  }

  // Render Detailed Comparison Table
  renderMultiCompanyTable(filteredYears, compConfigs.filter(c => enabled.includes(c.key)), metricKey, metricInfo);
}

// Side-by-Side Comparison Table
function renderMultiCompanyTable(years, activeCompanies, metricKey, metricInfo) {
  const container = document.getElementById('multi-comp-table-container');
  if (!container) return;

  let html = `<table class="data-table"><thead><tr>
    <th>Fiscal Year</th>`;
  
  activeCompanies.forEach(c => {
    html += `<th class="num" style="color:${c.color}">${c.name}</th>`;
  });
  html += '</tr></thead><tbody>';

  years.forEach(yr => {
    html += `<tr><td class="bold">${yr}</td>`;
    activeCompanies.forEach(c => {
      const rec = c.metrics.find(m => m.fiscal_year === yr);
      const val = rec ? rec[metricKey] : null;
      let displayVal = '—';
      if (val !== null && val !== undefined) {
        displayVal = metricInfo.format(val);
      }
      html += `<td class="num">${displayVal}</td>`;
    });
    html += '</tr>';
  });

  // Calculate Average Row
  html += `<tr class="total-row"><td class="bold">10-Yr Average / Total</td>`;
  activeCompanies.forEach(c => {
    const validVals = years
      .map(yr => {
        const rec = c.metrics.find(m => m.fiscal_year === yr);
        return rec ? rec[metricKey] : null;
      })
      .filter(v => v !== null && v !== undefined && !isNaN(v));

    let avgDisplay = '—';
    if (validVals.length > 0) {
      const avg = validVals.reduce((a, b) => a + b, 0) / validVals.length;
      avgDisplay = metricInfo.format(avg);
    }
    html += `<td class="num bold" style="color:${c.color}">${avgDisplay}</td>`;
  });
  html += '</tr></tbody></table>';

  if (['pn_ratio', 'net_sentiment_k', 'total_word_count', 'pos_words', 'neg_words'].includes(metricKey)) {
    html += '<p style="color:var(--text-muted); font-size:0.8rem; margin-top:0.75rem; font-style:italic;">* MD&A linguistic sentiment metrics are unavailable (—) for Microsoft Corporation (MSFT). Sentiment analysis is verified for primary research subjects Apple Inc. (AAPL) and Advanced Micro Devices (AMD).</p>';
  }

  container.innerHTML = html;
}

// 8. Original Cross-Company Dual Trajectory Chart
function renderCrossCompanyCharts() {
  const canvas = document.getElementById('chart-cross-company');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const aaplMetrics = state.data.comparison.metrics.AAPL;
  const amdMetrics = state.data.comparison.metrics.AMD;

  if (state.compChartInstance) {
    state.compChartInstance.destroy();
  }

  const datasets = [
    {
      label: 'Apple Indexed Stock Price (FY16=100)',
      data: aaplMetrics.map(m => m.indexed_price),
      borderColor: '#38bdf8',
      backgroundColor: '#38bdf8',
      borderWidth: 2.5,
      pointRadius: 4,
      tension: 0.1
    },
    {
      label: 'AMD Indexed Stock Price (FY16=100)',
      data: amdMetrics.map(m => m.indexed_price),
      borderColor: '#ef4444',
      backgroundColor: '#ef4444',
      borderWidth: 2.5,
      pointRadius: 4,
      tension: 0.1
    },
    {
      label: 'Apple MD&A P/N Ratio (x100 for scale)',
      data: aaplMetrics.map(m => m.pn_ratio * 100),
      borderColor: '#60a5fa',
      borderDash: [5, 5],
      pointRadius: 3,
      tension: 0.1
    },
    {
      label: 'AMD MD&A P/N Ratio (x100 for scale)',
      data: amdMetrics.map(m => m.pn_ratio * 100),
      borderColor: '#f87171',
      borderDash: [5, 5],
      pointRadius: 3,
      tension: 0.1
    }
  ];

  if (state.customCompany) {
    datasets.push({
      label: `${state.customCompany.ticker} Indexed Price (FY16=100)`,
      data: state.customCompany.metrics.map(m => m.indexed_price),
      borderColor: '#10b981',
      backgroundColor: '#10b981',
      borderWidth: 2.5,
      pointRadius: 4,
      tension: 0.1
    });
    const hasCustomSentiment = state.customCompany.metrics.some(m => m.pn_ratio !== null && m.pn_ratio !== undefined);
    if (hasCustomSentiment) {
      datasets.push({
        label: `${state.customCompany.ticker} MD&A P/N Ratio (x100 for scale)`,
        data: state.customCompany.metrics.map(m => m.pn_ratio ? m.pn_ratio * 100 : null),
        borderColor: '#34d399',
        borderDash: [5, 5],
        pointRadius: 3,
        tension: 0.1
      });
    }
  }

  state.compChartInstance = new Chart(ctx, {
    type: 'line',
    data: {
      labels: state.data.comparison.fiscal_years,
      datasets: datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: { color: '#94a3b8', font: { size: 12 }, padding: 15 }
        },
        tooltip: {
          backgroundColor: '#1e293b',
          titleColor: '#fff',
          bodyColor: '#cbd5e1',
          borderColor: '#334155',
          borderWidth: 1
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(255, 255, 255, 0.05)' },
          ticks: { color: '#94a3b8' }
        },
        y: {
          grid: { color: 'rgba(255, 255, 255, 0.05)' },
          ticks: {
            color: '#94a3b8',
            callback: (v) => v.toFixed(0)
          }
        }
      }
    }
  });
}

// 9. Third-Company Analysis Execution Logic
async function executeTickerAnalysis(ticker) {
  const cleanTicker = ticker.trim().toUpperCase();
  if (!cleanTicker.match(/^[A-Z]{1,5}$/)) {
    showAnalysisAlert(`Invalid ticker '${ticker}'. Ticker must contain 1 to 5 alphabetical characters (e.g., MSFT, NVDA, TSLA).`, 'warning');
    return;
  }

  hideAnalysisAlert();
  showProgressCard();
  updateProgress('Resolving SEC Central Index Key (CIK)...', 15, 1);

  // Check if live backend is reachable
  let analysisResult = null;

  try {
    const res = await fetch(`${state.backendUrl}/api/analyze`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ticker: cleanTicker })
    });

    if (res.ok) {
      updateProgress('Querying 10-K Filings & Financials...', 50, 3);
      analysisResult = await res.json();
      updateProgress('Processing MD&A Sentiment & Stock Prices...', 85, 4);
    } else {
      const errData = await res.json().catch(() => ({}));
      throw new Error(errData.error || `Server returned HTTP ${res.status}`);
    }
  } catch (backendErr) {
    console.log('[SEC Analyzer] Live backend request failed or unreachable:', backendErr.message);

    // Check if we have precomputed baseline for this ticker (e.g. MSFT)
    if (cleanTicker === 'MSFT' && state.data.msft_precomputed) {
      updateProgress('Loading verified 10-year MSFT dataset from precomputed audit baseline...', 75, 4);
      await new Promise(r => setTimeout(r, 600));
      analysisResult = state.data.msft_precomputed;
    } else {
      hideProgressCard();
      showAnalysisAlert(`
        <strong>Live SEC Backend Offline</strong><br>
        The backend API at <code>${state.backendUrl}</code> could not be reached.<br><br>
        <strong>How to run live analysis for any ticker:</strong><br>
        1. Open your terminal in the project directory.<br>
        2. Run: <code>python backend/server.py</code><br>
        3. Click "Analyze Company" again.<br><br>
        <em>Note: You can also click the <strong>"⚡ Load Verified MSFT Baseline"</strong> button to test the complete 3-company analysis feature immediately without starting a backend!</em>
      `, 'disclaimer');
      return;
    }
  }

  // Complete analysis
  updateProgress('Analysis successfully verified and loaded!', 100, 5);
  setTimeout(() => {
    hideProgressCard();
    setCustomCompany(analysisResult);
    showAnalysisSuccess(analysisResult);
  }, 500);
}

// Load Precomputed MSFT Baseline
function loadPrecomputedMSFT() {
  if (state.data.msft_precomputed) {
    setCustomCompany(state.data.msft_precomputed);
    showAnalysisSuccess(state.data.msft_precomputed);
  } else {
    // Attempt fetch
    fetch('data/msft_precomputed.json')
      .then(res => res.json())
      .then(data => {
        state.data.msft_precomputed = data;
        setCustomCompany(data);
        showAnalysisSuccess(data);
      })
      .catch(err => {
        showAnalysisAlert('Failed to load precomputed MSFT data: ' + err.message, 'warning');
      });
  }
}

// Set Third Company into Application
function setCustomCompany(analysisData) {
  state.customCompany = analysisData;
  const ticker = analysisData.ticker;

  // Add to PROFILES
  PROFILES.CUSTOM = {
    name: analysisData.company_name,
    ticker: ticker,
    cik: analysisData.cik,
    sector: 'Technology',
    industry: 'Enterprise Software & Cloud',
    fiscalEnd: 'June 30',
    exchange: 'NASDAQ',
    color: '#10b981',
    description: analysisData.profile?.description || `Analyzed SEC registrant ${analysisData.company_name} across 10 fiscal periods.`
  };

  // Update Header Button
  const btnCustom = document.getElementById('btn-custom-company');
  const labelCustom = document.getElementById('custom-company-label');
  if (btnCustom && labelCustom) {
    labelCustom.textContent = ticker === 'MSFT' ? 'Microsoft (MSFT)' : `${analysisData.company_name} (${ticker})`;
    btnCustom.classList.remove('hidden');
  }

  // Update Toolbar Checkbox
  const chipCustom = document.getElementById('chip-comp-custom');
  const labelToolbar = document.getElementById('label-comp-custom');
  const chkCustom = document.getElementById('chk-comp-custom');
  if (chipCustom && labelToolbar && chkCustom) {
    labelToolbar.textContent = `🏢 ${analysisData.company_name} (${ticker})`;
    chipCustom.classList.remove('hidden');
    chkCustom.checked = true;
  }

  // Update Compare Button Label
  const compLabel = document.getElementById('comp-btn-label');
  if (compLabel) {
    compLabel.textContent = `Three-Company Comparison`;
  }

  // Ensure active in multi-comp enabled list
  if (!state.multiComp.enabledCompanies.includes('CUSTOM')) {
    state.multiComp.enabledCompanies.push('CUSTOM');
  }

  renderAll();
}

// Remove Third Company Cleanly
function removeCustomCompany() {
  state.customCompany = null;
  delete PROFILES.CUSTOM;

  // Hide header button
  const btnCustom = document.getElementById('btn-custom-company');
  if (btnCustom) btnCustom.classList.add('hidden');

  // Hide toolbar checkbox
  const chipCustom = document.getElementById('chip-comp-custom');
  if (chipCustom) chipCustom.classList.add('hidden');

  // Reset compare button label
  const compLabel = document.getElementById('comp-btn-label');
  if (compLabel) compLabel.textContent = 'Compare AAPL vs AMD';

  // Remove from multi-comp enabled list
  state.multiComp.enabledCompanies = state.multiComp.enabledCompanies.filter(c => c !== 'CUSTOM');

  // If currently viewing CUSTOM, revert to AAPL
  if (state.currentCompany === 'CUSTOM') {
    state.currentCompany = 'AAPL';
    document.querySelectorAll('.company-btn').forEach(b => {
      b.classList.toggle('active', b.dataset.company === 'AAPL');
    });
  }

  // Hide success card in add-company section
  const successCard = document.getElementById('analyzed-summary-card');
  if (successCard) successCard.classList.add('hidden');

  renderAll();
}

// UI Progress Helpers
function showProgressCard() {
  const card = document.getElementById('analysis-progress-card');
  if (card) card.classList.remove('hidden');
}

function hideProgressCard() {
  const card = document.getElementById('analysis-progress-card');
  if (card) card.classList.add('hidden');
}

function updateProgress(msg, pct, stepNum) {
  const msgEl = document.getElementById('progress-status-msg');
  const pctEl = document.getElementById('progress-percent');
  const barEl = document.getElementById('progress-bar-fill');

  if (msgEl) msgEl.textContent = msg;
  if (pctEl) pctEl.textContent = `${pct}%`;
  if (barEl) barEl.style.width = `${pct}%`;

  for (let s = 1; s <= 5; s++) {
    const stepEl = document.getElementById(`step-${s}`);
    if (stepEl) {
      if (s < stepNum) {
        stepEl.className = 'step-item completed';
      } else if (s === stepNum) {
        stepEl.className = 'step-item active';
      } else {
        stepEl.className = 'step-item';
      }
    }
  }
}

function showAnalysisAlert(msg, type = 'disclaimer') {
  const alert = document.getElementById('analysis-alert');
  const title = document.getElementById('analysis-alert-title');
  const msgEl = document.getElementById('analysis-alert-msg');
  if (alert && title && msgEl) {
    title.textContent = type === 'warning' ? 'Notice' : 'Backend Status';
    msgEl.innerHTML = msg;
    alert.className = `callout ${type}`;
    alert.classList.remove('hidden');
  }
}

function hideAnalysisAlert() {
  const alert = document.getElementById('analysis-alert');
  if (alert) alert.classList.add('hidden');
}

function showAnalysisSuccess(data) {
  const card = document.getElementById('analyzed-summary-card');
  const title = document.getElementById('analyzed-company-title');
  const desc = document.getElementById('analyzed-company-desc');
  if (card && title && desc) {
    title.textContent = `✓ Successfully Analyzed ${data.company_name} (${data.ticker})`;
    const latest = data.metrics[data.metrics.length - 1];
    desc.innerHTML = `
      CIK: <strong>${data.cik}</strong> | Available Fiscal Periods: <strong>${data.fiscal_years.length} (${data.fiscal_years[0]} - ${data.fiscal_years[data.fiscal_years.length - 1]})</strong><br>
      FY2025 Revenue: <strong>${formatMillions(latest.net_sales_m)}</strong> | FY2025 Net Income: <strong>${formatMillions(latest.net_income_m)}</strong> | Net Profit Margin: <strong>${formatPercent(latest.net_profit_margin)}</strong> | MD&A P/N Ratio: <strong>${latest.pn_ratio ? latest.pn_ratio.toFixed(4) : '—'}</strong>
    `;
    card.classList.remove('hidden');
  }
}

// 10. Export Functions
function exportChartPNG() {
  if (!state.multiCompChartInstance) return;
  const imageURI = state.multiCompChartInstance.toBase64Image();
  const link = document.createElement('a');
  link.download = `SEC_Comparison_${state.multiComp.selectedMetric}_${new Date().toISOString().slice(0, 10)}.png`;
  link.href = imageURI;
  link.click();
}

function exportComparisonCSV() {
  const metricKey = state.multiComp.selectedMetric;
  const metricInfo = METRIC_DEFINITIONS[metricKey] || METRIC_DEFINITIONS.net_sales_m;
  const allYears = state.data.comparison.fiscal_years;
  const startIdx = allYears.indexOf(state.multiComp.startYear);
  const endIdx = allYears.indexOf(state.multiComp.endYear);
  const years = allYears.slice(Math.max(0, startIdx), endIdx >= startIdx ? endIdx + 1 : allYears.length);

  const activeCompanies = [
    { name: 'Apple Inc.', ticker: 'AAPL', metrics: state.data.comparison.metrics.AAPL },
    { name: 'Advanced Micro Devices', ticker: 'AMD', metrics: state.data.comparison.metrics.AMD }
  ];

  if (state.customCompany) {
    activeCompanies.push({
      name: state.customCompany.company_name,
      ticker: state.customCompany.ticker,
      metrics: state.customCompany.metrics
    });
  }

  const enabled = activeCompanies.filter(c => state.multiComp.enabledCompanies.includes(c.ticker === state.customCompany?.ticker ? 'CUSTOM' : c.ticker));

  let csv = `"Fiscal Year"`;
  enabled.forEach(c => {
    csv += `,"${c.name} (${c.ticker}) - ${metricInfo.name}"`;
  });
  csv += `\r\n`;

  years.forEach(yr => {
    csv += `"${yr}"`;
    enabled.forEach(c => {
      const rec = c.metrics.find(m => m.fiscal_year === yr);
      const val = rec ? rec[metricKey] : '';
      csv += `,${val !== null && val !== undefined ? val : ''}`;
    });
    csv += `\r\n`;
  });

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.download = `SEC_Analytics_${metricKey}_Comparison.csv`;
  link.href = url;
  link.click();
  URL.revokeObjectURL(url);
}

// Start application on DOM Content Loaded
document.addEventListener('DOMContentLoaded', initApp);
