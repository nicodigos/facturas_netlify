export const state = {
  config: null,
  graphToken: sessionStorage.getItem("graphToken") || "",
  driveId: "",
  databaseRows: [],
  databaseEtag: null,
  filteredRows: [],
  filters: {},
  pagination: {
    page: 1,
    pageSize: 10,
  },
  processed: {
    summaryRows: [],
    rawRows: [],
    pendingUploads: [],
    saved: false,
  },
};
