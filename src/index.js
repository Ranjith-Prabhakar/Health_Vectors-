const XLSX = require("xlsx");
const axios = require("axios");
const fs = require("fs");
const path = require("path");

const workbook = XLSX.readFile(
  path.join(__dirname, "assets", "input_excel_file_v1.xlsx")
);
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const rows = XLSX.utils.sheet_to_json(sheet, { defval: "NULL", raw: false });

const customerMap = new Map();

rows.forEach((row) => {
  const customerId = String(row["Customer ID"]);
  const age = String(row["age"]);
  const panelCode = String(row["panel_code"]);
  const panelName = String(row["panel_name"]);

  const parameter = {
    parameterName: String(row["Parameter Name"]),
    unit: String(row["Units"]),
    parameterCode: String(row["Parameter Code"]),
    value: String(row["Result"]),
    lowerRange: String(row["Low Range"]),
    upperRange: String(row["High Range"]),
    displayRange:
      row["Low Range"] === "NULL" && row["High Range"] === "NULL"
        ? "_"
        : `${String(row["Low Range"])}-${String(row["High Range"])}`,
  };

  // 1. Get or initialize the customer object
  let customer = customerMap.get(customerId);
  if (!customer) {
    customer = {
      "Customer ID": customerId,
      age: age,
      panelList: [],
    };
    customerMap.set(customerId, customer);
  }

  // 2. Find if the panel already exists within the customer's panelList
  let panel = customer.panelList.find((p) => p.panel_code === panelCode);

  // 3. If the panel doesn't exist, create it and add it to the panelList
  if (!panel) {
    panel = {
      panel_code: panelCode,
      panel_name: panelName,
      parameters: [],
    };
    customer.panelList.push(panel);
  }

  // 4. Add the current parameter to the found or newly created panel
  panel.parameters.push(parameter);
});

console.log("customerMap", customerMap.toString());
const finalJson = Array.from(customerMap.values());
fs.writeFileSync(
  path.join(__dirname, "assets", "result.json"),
  JSON.stringify(finalJson, null, 2)
);
console.log("✅ Data successfully written to result.json");

axios
  .post(
    "https://stage.myhealthvectors.com/testserver/receive-report",
    finalJson
  )
  .then((res) => {
    console.log("✅ Success:", res.data);
  })
  .catch((err) => {
    console.error("❌ Failed:", err.response?.status, err.response?.data);
  });
