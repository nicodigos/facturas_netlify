const COMPANY_OPTIONS = [
  "1001298527 ONTARIO INC",
  "10342548 CANADA INC",
  "10696480 CANADA LTD",
  "12433087 CANADA INC-MASTER",
  "13037622 CANADA INC",
  "9359-6633 QUEBEC INC",
  "9390-9216 QUEBEC INC",
  "D-TECH CONSTRUCTION",
  "TAYANTI-CANADA",
];

const BANK_OPTIONS = ["Scotiabank", "Desjardins", "National Bank"];

exports.handler = async function handler() {
  return {
    statusCode: 200,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      companyOptions: COMPANY_OPTIONS,
      bankOptions: BANK_OPTIONS,
      spHostname: process.env.SP_HOSTNAME || "",
      spSitePath: process.env.SP_SITE_PATH || "",
      spDriveName: process.env.SP_DRIVE_NAME || "Documents",
      receiptsDatabaseDir: process.env.RECEIPTS_DATABASE_DIR || "General/Sales receipts database",
      receiptsDatabaseCsv: "sales_receipts_database.csv",
    }),
  };
};
