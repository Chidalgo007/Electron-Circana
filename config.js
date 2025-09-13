const StoreModule = require("electron-store");
const Store = StoreModule.default || StoreModule; // handle default export

const store = new Store({
  name: "CircanaDashboard-config", // saved as CircanaDashboard-config.json
  defaults: {
    downloadPath: "",
    destinationPath: "",
    excelPath: "",
    npdPath: "",
    schedule: [],
  },
});

module.exports = store;
