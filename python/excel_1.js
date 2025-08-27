// # this will be a .exe file

// print("Excel: starting processing...")

// # example ... need to implement actual logic
// # df = pd.DataFrame({"Amount": [10, 20, 30]})
// # df["NewCol"] = df["Amount"] * 2
// # df.to_excel("output.xlsx", index=False)

// make the route of the files generic so any user can use the program.
const path = require("path");
const os = require("os");

const userHome = os.homedir();

const flatFilesPath = path.join(
  userHome,
  "NESTLE",
  "Commercial Development - Documents",
  "General",
  "03 Shopper Centricity",
  "Circana",
  "Flat Files"
);

console.log(flatFilesPath);

print("Excel: finished successfully");
