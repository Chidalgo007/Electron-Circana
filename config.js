import Store from "electron-store";

const store = new Store({
  name: "CircanaDashboard-config", // saved as CircanaDashboard-config.json
  defaults: {
    downloadPath: "",
    destinationPath: "",
    excelPath: "",
  },
});

export default store;
