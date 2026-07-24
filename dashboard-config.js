(function () {
  const localRefreshEndpoint = "http://127.0.0.1:8787/api/refresh-after-sales";
  const cloudRefreshEndpoint = "";
  const legacyCloudRefreshEndpoint = "https://fjdynamics.feishuapp.com/app/app_17aftyrd4rf/api/refresh-after-sales";
  const endpointKey = "sveaDashboardRefreshEndpoint";
  const tokenKey = "sveaDashboardRefreshToken";
  const params = new URLSearchParams(window.location.search);

  const explicitEndpoint = params.get("refreshEndpoint");
  if (explicitEndpoint) {
    window.localStorage.setItem(endpointKey, explicitEndpoint);
  }
  if (params.get("refreshApi") === "local") {
    window.localStorage.setItem(endpointKey, localRefreshEndpoint);
  }
  if (params.get("refreshApi") === "cloud") {
    if (cloudRefreshEndpoint) {
      window.localStorage.setItem(endpointKey, cloudRefreshEndpoint);
    } else {
      window.localStorage.removeItem(endpointKey);
    }
  }
  if (params.get("refreshApi") === "off") {
    window.localStorage.removeItem(endpointKey);
  }

  let savedEndpoint = window.localStorage.getItem(endpointKey) || "";
  if (savedEndpoint === legacyCloudRefreshEndpoint && !cloudRefreshEndpoint) {
    window.localStorage.removeItem(endpointKey);
    savedEndpoint = "";
  }
  const localPage = window.location.protocol === "file:" || ["127.0.0.1", "localhost"].includes(window.location.hostname);

  window.dashboardRefreshConfig = {
    afterSalesEndpoint: savedEndpoint || (localPage ? localRefreshEndpoint : ""),
    refreshToken: window.localStorage.getItem(tokenKey) || "",
  };
})();
