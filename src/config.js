// Client-specific settings for the Saint Eyewear month-end report.
export const NAVY = "#1F3A5F", GREEN = "#2E9E5B", RED = "#C0392B", ORANGE = "#DE8319";
export const CAT5 = ["#35609A", "#DE8319", "#0E9C8C", "#9B51B0", "#8E9AA6"];
export const INK = "#1a1a1a", MUTED = "#5a5a5a", GRID = "#DDE1E6";
export const CHART_DPI = 190;

// Allowlist of P&L section headings: an unknown heading stops the run.
export const PL_SECTIONS = {
  "trading income": "income", "income": "income", "revenue": "income",
  "cost of sales": "cogs", "cost of goods sold": "cogs", "direct costs": "cogs",
  "operating expenses": "opex", "expenses": "opex",
};
export const PL_DERIVED_ROWS = ["gross profit", "net profit", "net income", "operating profit"];

export const BS_SECTIONS = {
  "assets": null, "liabilities": null,
  "bank": "Bank", "current assets": "Current Assets", "fixed assets": "Fixed Assets",
  "non-current assets": "Fixed Assets",
  "current liabilities": "Current Liabilities",
  "non-current liabilities": "Non-Current Liabilities",
  "equity": "Equity",
};
export const LIABILITY_SHOWN_AS_BANK = ["Business Trans Acct", "Revolut USD Main"];
export const NET_ASSET_AGAINST_LIABILITY = ["Credit account (SAINT EYEWEAR LLC)"];
export const MANUAL_BANK_ACCOUNTS = ["CommBank USD"];
export const SUSPENSE_ACCOUNT = "Suspense";

export const TOL_EXACT = 0.01;
export const TOL_DRS_ABS = 250.0;
export const TOL_DRS_PCT = 0.005;
export const PRODUCT_COVERAGE_MIN = 0.99;
export const RESTATE_LINE_MIN = 1.0;

export const EXPENDITURE_EXCLUDE_PREFIXES = ["social media ads", "rounding"];
export const FEE_ACCOUNTS = { Shopify: "Fees Shopify", Paypal: "Fees Paypal", Afterpay: "Fees AfterPay" };
export const DRS_AD_MAP = [
  ["advertisment spend - meta", "Social Media Ads - Meta", 1.0],
  ["advertisment spend - google", "Social Media Ads - Google", 1.1],
  ["advertisment spend - pinterest", "Social Media Ads - Pinterest", 1.0],
  ["advertisment spend - snap", "Social Media Ads - Snapchat", 1.1],
  ["advertisment spend - tiktok", "Social Media Ads - Tiktok", 1.0],
];
