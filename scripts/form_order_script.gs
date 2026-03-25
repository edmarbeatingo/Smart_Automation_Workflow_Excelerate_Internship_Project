function batchProcessOrders() {

  const ss = SpreadsheetApp.getActive();
  const source = ss.getSheetByName("Form Responses 1");
  const master = ss.getSheetByName("PRODUCT_MASTER");
  const orderSheet = ss.getSheetByName("ORDER_FORM (Responses)");

  if (!source || !master || !orderSheet) {
    Logger.log("Missing required sheet.");
    return;
  }

  // CLEAR ORDER SHEET (SYNC MODE)
  const existingLastRow = orderSheet.getLastRow();
  if (existingLastRow > 1) {
    orderSheet
      .getRange(2, 1, existingLastRow - 1, orderSheet.getLastColumn())
      .clearContent();
  }

  const territoryMap = {
    "Mumbai": { name: "Rajesh Sharma", prefix: "MUM" },
    "Delhi": { name: "Priya Mehta", prefix: "DEL" },
    "Bengaluru": { name: "Arjun Rao", prefix: "BLR" },
    "Chennai": { name: "Kavya Iyer", prefix: "CHN" },
    "Hyderabad": { name: "Vikram Reddy", prefix: "HYD" }
  };

  const headers = source.getRange(1,1,1,source.getLastColumn()).getValues()[0];
  const lastRow = source.getLastRow();
  if (lastRow < 2) return;

  const data = source.getRange(2,1,lastRow-1,headers.length).getValues();
  const masterData = master.getRange(2,1,master.getLastRow()-1,6).getValues();

  // RESET SEQUENCE (SYNC MODE)
  let maxSeq = 0;

  data.forEach(row => {

    let timestamp = "";
    let customerName = "";
    let contactNo = "";
    let customerEmail = "";
    let location = "";

    headers.forEach((header, i) => {
      const headerLower = header.toString().toLowerCase();

      if (headerLower.includes("timestamp")) {
        timestamp = row[i];
      }
      if (headerLower.includes("customer name")) {
        customerName = row[i];
      }
      if (headerLower.includes("contact")) {
        contactNo = row[i];
      }
      if (headerLower.includes("email")) {
        customerEmail = row[i];
      }
      if (headerLower.includes("location")) {
        location = (row[i] || "").toString().trim();
      }
    });

    if (!timestamp) return;

    const territory = territoryMap[location];
    if (!territory) {
      Logger.log("Unknown territory: " + location);
      return;
    }

    const orderDate = Utilities.formatDate(
      new Date(timestamp),
      Session.getScriptTimeZone(),
      "M/d/yyyy"
    );

    headers.forEach((header, index) => {

      if (!header.includes("Specifications")) return;

      const spec = row[index];
      if (!spec) return;

      let qty = 0;
      for (let i = index + 1; i < headers.length; i++) {
        if (headers[i].toLowerCase().includes("quantity")) {
          qty = parseFloat(row[i]) || 0;
          break;
        }
      }

      if (qty <= 0) return;

      const match = masterData.find(r =>
        r[3] &&
        r[3].toString().trim() === spec.toString().trim()
      );

      if (!match) {
        Logger.log("No PRODUCT_MASTER match for: " + spec);
        return;
      }

      const productID = match[0];
      const productName = match[1];
      const category = match[2];
      const uom = match[4];
      const unitPrice = parseFloat(match[5]) || 0;
      const orderAmount = unitPrice * qty;

      // Increment sequence during rebuild
      maxSeq++;
      const paddedSeq = Utilities.formatString("%05d", maxSeq);

      const datePart = Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "MMddyyyy"
      );

      const orderID = `${territory.prefix}-${datePart}-FO-${paddedSeq}`;

      orderSheet.appendRow([
        orderID,
        orderDate,
        territory.name,
        productID,
        productName,
        category,
        spec,
        uom,
        unitPrice,
        location,
        customerName,
        contactNo,
        customerEmail,
        qty,
        orderAmount,
        Utilities.formatDate(
          new Date(timestamp),
          Session.getScriptTimeZone(),
          "M/d/yyyy HH:mm:ss"
        )
      ]);

    });

  });

}
