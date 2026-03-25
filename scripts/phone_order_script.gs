function onEdit(e) {

  if (!e) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (["PRODUCT_MASTER", "CONTROL", "MASTERFILE"].includes(sheetName)) return;

  const row = e.range.getRow();
  if (row < 2) return;

  const master = ss.getSheetByName("PRODUCT_MASTER");
  if (!master) return;

  const masterData = master.getDataRange().getValues();
  masterData.shift();


  const orderIdCol = 1;
  const orderDateCol = 2;
  const salespersonCol = 3;
  const productIdCol = 4;
  const productNameCol = 5;
  const categoryCol = 6;
  const specsCol = 7;
  const uomCol = 8;
  const priceCol = 9;
  const locationCol = 10;
  const customerCol = 11;
  const contactCol = 12;
  const emailCol = 13;
  const quantityCol = 14;
  const amountCol = 15;
  const timestampCol = 16;

  const territoryMap = {
    "Rajesh Sharma": { prefix: "MUM", location: "Mumbai" },
    "Priya Mehta": { prefix: "DEL", location: "Delhi" },
    "Arjun Rao": { prefix: "BLR", location: "Bengaluru" },
    "Kavya Iyer": { prefix: "CHN", location: "Chennai" },
    "Vikram Reddy": { prefix: "HYD", location: "Hyderabad" }
  };

  if (!territoryMap[sheetName]) return;
  const territory = territoryMap[sheetName];

  const rowRange = sheet.getRange(row, 4, 1, 6);
  let [productId, productName, category, specs] = rowRange.getValues()[0];

  let filtered = masterData;

  if (productId) filtered = filtered.filter(r => r[0] === productId);
  if (productName) filtered = filtered.filter(r => r[1] === productName);
  if (category) filtered = filtered.filter(r => r[2] === category);
  if (specs) filtered = filtered.filter(r => r[3] === specs);

  if (filtered.length > 0) {
    rebuildDropdowns(sheet, row, filtered);

    if (filtered.length === 1 ||
        e.range.getColumn() === productIdCol ||
        e.range.getColumn() === specsCol) {
      fillAll(sheet, row, filtered[0]);
    }
  }

  const quantityRaw = sheet.getRange(row, quantityCol).getValue();
  const unitPriceRaw = sheet.getRange(row, priceCol).getValue();

  const cleanQty = parseFloat(quantityRaw);
  const cleanPrice = parseFloat(
    unitPriceRaw ? unitPriceRaw.toString().replace(/,/g, '') : 0
  );

  if (!isNaN(cleanQty) && !isNaN(cleanPrice) && cleanQty > 0) {

    const total = cleanQty * cleanPrice;

    sheet.getRange(row, amountCol)
         .setValue(total)
         .setNumberFormat("#,##0.00");

  } else {
    sheet.getRange(row, amountCol).clearContent();
  }

  const orderAmount = sheet.getRange(row, amountCol).getValue();
  const orderDate = sheet.getRange(row, orderDateCol).getValue();
  const customerName = sheet.getRange(row, customerCol).getValue();
  const contactNo = sheet.getRange(row, contactCol).getValue();
  const email = sheet.getRange(row, emailCol).getValue();

  if (orderDate &&
      customerName &&
      contactNo &&
      email &&
      cleanQty > 0 &&
      orderAmount > 0) {

    sheet.getRange(row, salespersonCol).setValue(sheetName);
    sheet.getRange(row, locationCol).setValue(territory.location);

  } else {

    sheet.getRange(row, salespersonCol).clearContent();
    sheet.getRange(row, locationCol).clearContent();
  }


  let currentOrderId = sheet.getRange(row, orderIdCol).getValue();

  if (!currentOrderId &&
      orderDate &&
      customerName &&
      contactNo &&
      email &&
      cleanQty > 0 &&
      orderAmount > 0) {

    const masterFile = ss.getSheetByName("MASTERFILE");
    const activeCount = masterFile.getLastRow() >= 2
      ? masterFile.getLastRow() - 1
      : 0;

    const nextSequence = activeCount + 1;

    const dateObj = new Date(orderDate);
    const month = String(dateObj.getMonth() + 1).padStart(2, '0');
    const day = String(dateObj.getDate()).padStart(2, '0');
    const year = dateObj.getFullYear();
    const datePrefix = month + day + year;

    const formattedCounter = String(nextSequence).padStart(5, '0');

    currentOrderId =
      territory.prefix + "-" + datePrefix + "-PO-" + formattedCounter;

    sheet.getRange(row, orderIdCol).setValue(currentOrderId);

    sheet.getRange(row, timestampCol)
         .setValue(new Date())
         .setNumberFormat("MM/dd/yyyy HH:mm:ss");
  }

  reconcileMasterFile();
}


function onChange(e) {

  if (!e) return;

  if (["REMOVE_ROW", "REMOVE_GRID", "EDIT"].includes(e.changeType)) {
    reconcileMasterFile();
  }
}

function reconcileMasterFile() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterFile = ss.getSheetByName("MASTERFILE");
  const controlSheet = ss.getSheetByName("CONTROL");

  const salespersonSheets = [
    "Rajesh Sharma",
    "Priya Mehta",
    "Arjun Rao",
    "Kavya Iyer",
    "Vikram Reddy"
  ];

  let validOrders = [];

  salespersonSheets.forEach(name => {

    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();

    data.forEach(row => {

      const orderId = row[0];
      const orderDate = row[1];
      const customer = row[10];
      const contact = row[11];
      const email = row[12];
      const qty = row[13];
      const amount = row[14];

      if (orderId &&
          orderDate &&
          customer &&
          contact &&
          email &&
          qty > 0 &&
          amount > 0) {

        validOrders.push(row);
      }
    });

  });

  if (masterFile.getLastRow() > 1) {
    masterFile.getRange(2, 1, masterFile.getLastRow() - 1, 16).clearContent();
  }

  if (validOrders.length > 0) {
    masterFile.getRange(2, 1, validOrders.length, 16)
              .setValues(validOrders);
  }

  controlSheet.getRange("A2").setValue(validOrders.length);
}


function rebuildDropdowns(sheet, row, filtered) {

  const ids = [...new Set(filtered.map(r => r[0]))];
  const names = [...new Set(filtered.map(r => r[1]))];
  const categories = [...new Set(filtered.map(r => r[2]))];
  const specs = [...new Set(filtered.map(r => r[3]))];

  setValidation(sheet, row, 4, ids);
  setValidation(sheet, row, 5, names);
  setValidation(sheet, row, 6, categories);
  setValidation(sheet, row, 7, specs);
}

function fillAll(sheet, row, match) {

  sheet.getRange(row, 4).setValue(match[0]);
  sheet.getRange(row, 5).setValue(match[1]);
  sheet.getRange(row, 6).setValue(match[2]);
  sheet.getRange(row, 7).setValue(match[3]);
  sheet.getRange(row, 8).setValue(match[4]);
  sheet.getRange(row, 9).setValue(match[5]);
}

function setValidation(sheet, row, col, values) {

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(row, col).setDataValidation(rule);
}
