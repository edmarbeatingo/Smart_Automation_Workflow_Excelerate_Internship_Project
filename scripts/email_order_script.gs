function extractEmailOrders() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName("Email_Orders");
  const masterSheet = ss.getSheetByName("PRODUCT_MASTER");

  if (!orderSheet) throw new Error("Email_Orders sheet not found");
  if (!masterSheet) throw new Error("PRODUCT_MASTER sheet not found");

  const masterData = masterSheet
    .getRange(2, 1, masterSheet.getLastRow() - 1, 6)
    .getValues();

  const territoryMap = {
    "Mumbai": { name: "Rajesh Sharma", prefix: "MUM" },
    "Delhi": { name: "Priya Mehta", prefix: "DEL" },
    "Bengaluru": { name: "Arjun Rao", prefix: "BLR" },
    "Chennai": { name: "Kavya Iyer", prefix: "CHN" },
    "Hyderabad": { name: "Vikram Reddy", prefix: "HYD" }
  };

  const threads = GmailApp.search('is:unread subject:"Sales Order -"');
  Logger.log("Threads found: " + threads.length);

  threads.forEach(thread => {

    thread.getMessages().forEach(message => {

      const bodyText = message.getPlainBody();

      /***********************
       SIMPLE FIELD EXTRACTION
      ************************/
      function getField(field) {
        const regex = new RegExp(field + "\\s*:\\s*(.*)", "i");
        const match = bodyText.match(regex);
        return match ? match[1].trim() : "";
      }

      const orderDateStr = getField("Order Date");
      const customerName = getField("Customer Name");
      const email = getField("Customer Email");
      const location = getField("Location");
      const quantity = Number(getField("Quantity"));

      if (!orderDateStr || !customerName || !email || !location || quantity <= 0) {
        Logger.log("Missing required order fields.");
        return;
      }

      /***********************
       PHONE EXTRACTION (STRICT)
      ************************/
      let contactNo = "";
      const contactMatch = bodyText.match(/Customer Contact No\s*:\s*([^\n\r]+)/i);

      if (contactMatch) {
        contactNo = contactMatch[1].trim();
      }

      contactNo = contactNo.replace(/\s+/g, "");

      const validFormat = /^\+\d{1,3}-\d{6,15}$/;

      if (!validFormat.test(contactNo)) {
        Logger.log("Invalid phone format: " + contactNo);
        return;
      }

      /***********************
       SAFE DATE PARSING
      ************************/
      if (!orderDateStr.includes("-")) {
        Logger.log("Invalid Order Date: " + orderDateStr);
        return;
      }

      const parts = orderDateStr.split("-");

      if (parts.length !== 3) {
        Logger.log("Malformed Order Date: " + orderDateStr);
        return;
      }

      const dateObj = new Date(
        parseInt(parts[2], 10),
        parseInt(parts[0], 10) - 1,
        parseInt(parts[1], 10)
      );

      if (isNaN(dateObj.getTime())) {
        Logger.log("Date conversion failed: " + orderDateStr);
        return;
      }

      const datePrefix = Utilities.formatDate(
        dateObj,
        Session.getScriptTimeZone(),
        "MMddyyyy"
      );

      /***********************
       SPECIFICATIONS EXTRACTION
      ************************/
      let specs = "";

      const specStart = bodyText.indexOf("Specifications:");
      const qtyStart = bodyText.indexOf("Quantity:");

      if (specStart !== -1 && qtyStart !== -1) {
        specs = bodyText
          .substring(specStart + "Specifications:".length, qtyStart)
          .replace(/\n/g, " ")
          .replace(/\r/g, " ")
          .replace(/\s+/g, " ")
          .trim();
      }

      if (!specs) {
        Logger.log("Specifications missing.");
        return;
      }

      const normalizedSpecs = specs
        .replace(/\s+/g, " ")
        .trim()
        .toLowerCase();

      const product = masterData.find(row =>
        row[3].toString().replace(/\s+/g, " ").trim().toLowerCase() === normalizedSpecs
      );

      if (!product) {
        Logger.log("Product not found for specs: " + specs);
        return;
      }

      const territory = territoryMap[location];
      if (!territory) {
        Logger.log("Invalid location: " + location);
        return;
      }

      const productId = product[0];
      const productName = product[1];
      const category = product[2];
      const uom = product[4];

      /***********************
       SAFE PRICE PARSING
      ************************/
      const rawPrice = product[5]
        .toString()
        .replace(/,/g, "")
        .replace(/[^\d.]/g, "");

      const price = parseFloat(rawPrice);

      if (isNaN(price)) {
        Logger.log("Invalid price in PRODUCT_MASTER: " + product[5]);
        return;
      }

      const amount = quantity * price;

      /***********************
       GLOBAL CONTINUOUS EO SEQUENCE
      ************************/
      let existingIds = [];
      const lastRow = orderSheet.getLastRow();

      if (lastRow > 1) {
        existingIds = orderSheet
          .getRange(2, 1, lastRow - 1, 1)
          .getValues()
          .flat()
          .filter(String);
      }

      let maxSeq = 0;

      existingIds.forEach(id => {
        if (id.includes("-EO-")) {
          const seq = parseInt(id.split("-EO-")[1], 10);
          if (!isNaN(seq) && seq > maxSeq) {
            maxSeq = seq;
          }
        }
      });

      const nextSeq = maxSeq + 1;
      const formattedSeq = String(nextSeq).padStart(5, "0");

      const orderId =
        territory.prefix + "-" + datePrefix + "-EO-" + formattedSeq;

      const completionTime = Utilities.formatDate(
        message.getDate(),
        Session.getScriptTimeZone(),
        "M/d/yyyy HH:mm:ss"
      );

      /***********************
       APPEND CLEAN ROW
      ************************/
      const newRow = [
        orderId,
        dateObj,
        territory.name,
        productId,
        productName,
        category,
        specs,
        uom,
        price,
        location,
        customerName,
        contactNo,
        email,
        quantity,
        amount,
        completionTime
      ];

      orderSheet.appendRow(newRow);
      message.markRead();

      Logger.log("Order added: " + orderId);

    });

  });

}
