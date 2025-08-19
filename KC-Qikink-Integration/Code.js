function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // your existing menu items…
  let menu = ui.createMenu('🛠️ KC Tools')
    .addItem("Update Statuses", "processOrdersForStatusUpdates") // Menu item and function name

  menu = kcRegisterGstCsvMenuItem(menu);
  menu.addToUi();
}

function myFunction() {
  SpreadsheetApp.getActiveSpreadsheet().toast("KC Tools -> Update Status Clicked", "Info");
}

function callOpenApi(url, method, headers, body) {
  var options = {
    'method': method,
    'headers': headers,
    'payload': body
  };

  // Logger.log("***** Call URL" + url)
  // Logger.log("***** Call Options")
  // Logger.log(options)
  var response = UrlFetchApp.fetch(url, options);
  Logger.log("***** Call Response", response)
  var content = response.getContentText();
  var json = JSON.parse(content);
  return json;
}

function getCellValue(sheetName, cellAddress) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  var cellValue = sheet.getRange(cellAddress).getValue();
  // Logger.log("In Sheet '"+sheetName+"', Value at '"+cellAddress+"' = "+ cellValue);
  return cellValue;
}

function callQikinkSandboxAPI(endpoint, method, headers, body) {
  //getting the Sandbox URL from googlesheet. It is present in "Automation" Sheet
  const sandboxURL = getCellValue("Automation", "J2");
  const urlEndpoint = sandboxURL + endpoint;
  console.log("URL Enpoint", urlEndpoint);
  return callOpenApi(urlEndpoint, method, headers, body)
}

function callQikinkLiveAPI(endpoint, method, headers, body) {
  //getting the Live URL from googlesheet. It is present in "Automation" Sheet
  const liveURL = getCellValue("Automation", "J3");
  const urlEndpoint = liveURL + endpoint;
  console.log("URL Enpoint", urlEndpoint);
  return callOpenApi(urlEndpoint, method, headers, body)
}

function getAuthorizationToken(isLive) {
  var endpoint = "api/token";
  var method = "POST";

  var headers = {
    'Content-Type': 'application/x-www-form-urlencoded'
  };

  if (!isLive) { //Sandbox API    
    var clientid = getCellValue("Automation", "H2");
    var clientsecret = getCellValue("Automation", "I2");
    var body = {
      'ClientId': clientid,
      'client_secret': clientsecret
    };
    var result = callQikinkSandboxAPI(endpoint, method, headers, body);
    Logger.log(result);
    return result
  } else { //Live API
    var clientid = getCellValue("Automation", "H3");
    var clientsecret = getCellValue("Automation", "I3");
    var body = {
      'ClientId': clientid,
      'client_secret': clientsecret
    };
    var result = callQikinkLiveAPI(endpoint, method, headers, body);
    Logger.log(result);
    return result
  }
}

function fetchOrders(pgNo) {
  const isLive = true;
  var endpoint = "api/order?page_no=" + (pgNo || "1");
  var method = "GET";
  const authResult = getAuthorizationToken(isLive);
  if (!isLive) {
    var clientid = getCellValue("Automation", "H2");
    var headers = {
      'ClientId': clientid,
      'Accesstoken': authResult.Accesstoken,
    };
    var body = {

    };
    var result = callQikinkSandboxAPI(endpoint, method, headers, body);
    Logger.log(result);
  } else {
    var clientid = getCellValue("Automation", "H3");
    var headers = {
      'ClientId': clientid,
      'Accesstoken': authResult.Accesstoken,
    };
    var body = {
    };
    var result = callQikinkLiveAPI(endpoint, method, headers, body);
    // Logger.log(result);
    return result;
  }

}

function fetchQikinkOrderByID(orderNo) {
  const isLive = true;
  var endpoint = "api/order?id=" + orderNo;
  var method = "GET";
  const authResult = getAuthorizationToken(isLive);
  if (!isLive) {
    var clientid = getCellValue("Automation", "H2");
    var headers = {
      'ClientId': clientid,
      'Accesstoken': authResult.Accesstoken,
    };
    var body={};
    var result = callQikinkSandboxAPI(endpoint, method, headers, body);
    // Logger.log(result);
    return result;
  } else {
    var clientid = getCellValue("Automation", "H3");
    var headers = {
      'ClientId': clientid,
      'Accesstoken': authResult.Accesstoken,
    };
    var body={};
    var result = callQikinkLiveAPI(endpoint, method, headers, body);
    // Logger.log(result);
    return result;
  }

}

function createOrder() {
  const isLive = false;
  var endpoint = "api/order/create";
  var method = "POST";
  const authResult = getAuthorizationToken(isLive);

  var payload1 = {
    "order_number": "ORD04e115",
    "qikink_shipping": "1",
    "gateway": "COD",
    "total_order_value": "500.00",
    "line_items": [
      {
        "search_from_my_products": 0,
        "quantity": "1",
        "print_type_id": 1,
        "price": "500",
        "sku": "MRnHs-Pu-S",
        "designs": [
          {
            "design_code": "Floral_A",
            "width_inches": "10",
            "height_inches": "12",
            "placement_sku": "fr",
            "design_link": "https://kidlingoo.com/flowers-name-in-english/",
            "mockup_link": "https://kidlingoo.com/flowers-name-in-english/"
          }
        ]
      }
    ],
    "add_ons": [
      {
        "box_packing": 0,
        "gift_wrap": 0,
        "rush_order": 0,
        "custom_letter": 0
      }
    ],
    "shipping_address": {
      "first_name": "first_name",
      "last_name": "last_name",
      "address1": "adrress_1...",
      "phone": "9876543210",
      "email": "sample@gmail.com",
      "city": "coimbatore",
      "zip": "641004",
      "province": "ABC",
      "country_code": "IN"
    }
  };

  var payload = {

    "order_number": "ORD04e14",

    "qikink_shipping": "1",

    "gateway": "COD",

    "total_order_value": "500.00",

    "line_items": [

      {

        "search_from_my_products": 0,

        "quantity": "1",

        "print_type_id": 1,

        "price": "500",

        "sku": "MRnHs-Pu-S",

        "designs": [

          {

            "design_code": "Floral_A",

            "width_inches": "10",

            "height_inches": "12",

            "placement_sku": "fr",

            "design_link": "https://kidlingoo.com/flowers-name-in-english/",

            "mockup_link": "https://kidlingoo.com/flowers-name-in-english/"

          }

        ]

      }

    ],

    "add_ons": [

      {

        "box_packing": 0,

        "gift_wrap": 0,

        "rush_order": 0,

        "custom_letter": 0

      }

    ],

    "shipping_address": {

      "first_name": "aaa",

      "last_name": "bb",

      "address1": "erdetgfhfhhghhg",

      "phone": "1234567890",

      "email": "aaa@gmail.com",

      "city": "xx",

      "zip": "641659",

      "province": "tamilnadu",

      "country_code": "IN"

    }

  }

  if (!isLive) {
    var clientid = getCellValue("Automation", "H2");
    var headers = {
      'ClientId': clientid,
      'Accesstoken': authResult.Accesstoken,
    };
    var result = callQikinkSandboxAPI(endpoint, method, headers, payload);
    Logger.log(result);
  } else {
    var clientid = getCellValue("Automation", "H3");
    var headers = {
      'ClientId': clientid,
      'Accesstoken': authResult.Accesstoken,
    };
    var result = callQikinkLiveAPI(endpoint, method, headers, payload);
    Logger.log(result);
  }

}

/**
 * Qikink Response Sample
 * {
    shipping_type=QikinkDomesticShipping,
    line_items=[
        {
            designs=[
                {
                    mockup_url=null,
                    height_inches=10.32,
                    design_code=105ThereisaGirlShecallsmePapa,
                    design_mockup_url=null,
                    placement=Front,
                    design_url=https: //sgp1.digitaloceanspaces.com/cdn.qikink.com/erp2/assets/designs/172086/1740207826.png,
                    width_inches=12.62,
                    printing_cost=100
                }
            ],
            sku=MRnHs-Bk-XL,
            quantity=1,
            price=160
        }
    ],
    number=172086_4848,
    status=Exception,
    total_order_value=899.00,
    created_on=2025-04-2301: 28: 11,
    payment_type=COD,
    live_date=2025-04-2309: 15: 10,
    order_id=1.15364424E10,
    shipping={
        phone=9625359442,
        country_code=IN,
        awb=80781000775,
        courier_provider_name=BluedartExpress,
        first_name=Shad,
        province=DL,
        city=SouthDelhi,
        zip=110025,
        tracking_link=https: //courierupdates.com/?awb=80781000775,
        last_name=Khan,
        email=shadw8826@gmail.com
    }
}
 */

/*
Plan
1. Run the program every day at 2 am and 2 pm
2 Create a new sheet 'Order Sync Logs' with below columns - Done
- - SyncLane = Qikink or Shopify
- - Timestamp = Current date and time
- - Order ID
- - ChangeType = NewOrder, StatusUpdate, MarkAsPaid, Return
- - Old State
- - New State
- - Notes
3 First it fills all the new orders from Qikink. - We wont be doing this automation because the api doesnot have all the fields, also the tax anyways needs to be added manually for prepaid orders. Also the Design drop down needs to be populated manually. So we will get all the details from qikink invoice reports and used it as it is done currently.
3.1 For this in the Qikink orders sheet, fetch the last order number - Done
3.2 See if there are any new orders using Qikink APIs
3.3 Insert new order rows in Qikink order sheet
3.4 Insert entries for each new order in Order Sync Logs sheet
4 Update the delivery status of pending orders
4.1 Get all order IDs from Qikink order sheet, where status is blank or Prepaid or Partially Paid
4.2 Check for each orders if any of it is showing Delivered in Qikink. 
4.2.1 If yes then change the status to "delivered" in sheet.
4.2.2 For delivered order, mark it as paid in shopify if the order is COD type.
4.2.3 For delivered order, add entry in Order Sync Logs sheet
4.3 Check for each orders if it is shipped and its tracking link is available 
4.3.1 If yes then update the tracking details in Google sheet
4.3.2 For trackings updates, add entry in Order Sync Logs sheet
4.3.3 Update shopify to add the fulfillment details
4.3.4 For fulfillment updates, add entry in Order Sync Logs sheet
4.4 Get all order IDs from Qikink order sheet and log entry in Order Sync Logs sheet if the status is Exception/RTO Initiated/Returned
*/

// 4 Update the delivery status of pending orders
function processOrdersForStatusUpdates() {
  SpreadsheetApp.getActiveSpreadsheet().toast("KC Tools -> Process Orders for Status Updates Started", "Info");
  const qikinkOrderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Qikink Orders');
  const data = qikinkOrderSheet.getDataRange().getValues(); // Get all data in the sheet
  const header = data[0];
  const statusIndex = header.indexOf("Status"); // Find the "status" column index  
  if (statusIndex === -1) {
    Logger.log("No 'Status' column found.");
    return;
  }
  const trackingColumnIndex = header.indexOf("TrackingLink"); // Find the "status" column index
  const awbColumnIndex = header.indexOf("AwbNo"); 
  const courierColumnIndex = header.indexOf("CourierPartner"); 
  const customerNameColumnIndex = header.indexOf("Customer Name"); 
  if (trackingColumnIndex === -1) {
    Logger.log("No 'TrackingLink' column found.");
    return;
  }

  //Fetch Qikink Orders for 80 Pages (1 page = 10 orders)
  var qikinkOrders = [];
  for (let pgno = 1; pgno <= 80; pgno++) {
    var limitCheck = pgno % 25;
    
    if (limitCheck == 0) {
      Utilities.sleep(5000); // 5000 milliseconds = 5 seconds wait to avoid rate limit
    }
    qikinkOrders.push(...fetchOrders(pgno));
  }
  Logger.log("Total Qikink Orders Fetched = " + qikinkOrders.length)

  for (let i = 1; i < data.length; i++) {
    const oldStatus = data[i][statusIndex].toString().trim().toLowerCase();
    const oldTrackingValue = data[i][trackingColumnIndex].toString().trim().toLowerCase();
    const orderNo = data[i][5].toString().trim().toLowerCase();

    // 4.1 Get all order IDs from Qikink order sheet, where status is blank or Prepaid or Partially Paid
    if (orderNo != "" && (oldStatus === "" || oldStatus === "prepaid" || oldStatus === "partially paid" || oldStatus === "cancelled")) {
      Logger.log(orderNo + "(" + (i + 1) + "," + (statusIndex + 1) + ")" + "::" + oldStatus)      
      const foundQikinkOrder = qikinkOrders.filter(order => order.number === orderNo)
      if (foundQikinkOrder && foundQikinkOrder.length > 0) {
        // 4.2 Check for each orders if any of it is showing Delivered in Qikink.
        if (foundQikinkOrder[0].status === "Delivered") {// this means we have found a delivered order
          // 4.2.1 If yes then change the status to "delivered" in sheet.
          const newStatus = 'Delivered'
          // qikinkOrderSheet.getRange(i + 1, statusIndex + 1).setValue(newStatus); // +1 because sheet rows/columns are 1-indexed
          Logger.log(orderNo + "(" + (i + 1) + "," + (statusIndex + 1) + ")" + "::" + oldStatus + "::" + newStatus)
          var logNotes = "SHEET -> changed status to '" + newStatus + "'; ";
          // 4.2.2 For delivered order, mark it as paid in shopify if the order is COD type.
          if (foundQikinkOrder[0].payment_type === "COD") {
            // logNotes += "SHOPIFY -> " + updateOrderStatusInShopify(convertQikinkOrderNumberToShopifyNumber(foundQikinkOrder[0].number), "success", null);
          }
          // 4.2.3 For delivered order, add entry in Order Sync Logs sheet
          addOrderSyncLog("Qikink", orderNo, "StatusUpdate", oldStatus, newStatus, logNotes)
        }
      }
      
      // 4.3 Check for each orders if it is shipped and its tracking link is available 
      const foundShippedQikinkOrder = qikinkOrders.filter(order => order.number === orderNo && order.shipping.awb != null)
      if (foundShippedQikinkOrder && foundShippedQikinkOrder.length > 0) {
        const newTrackingValue = foundShippedQikinkOrder[0].shipping.tracking_link
        const newAwbValue = foundQikinkOrder[0].shipping.awb
        const newCourierValue = foundQikinkOrder[0].shipping.courier_provider_name
        // Logger.log(foundShippedQikinkOrder[0]);
        // 4.3.1 If order is shipped and GS column is blank, then update the tracking details in Google sheet
        if(newTrackingValue!=null) {
          qikinkOrderSheet.getRange(i + 1, trackingColumnIndex + 1).setValue(newTrackingValue); // +1 because sheet rows/columns are 1-indexed
          qikinkOrderSheet.getRange(i + 1, awbColumnIndex + 1).setValue(newAwbValue); // +1 because sheet rows/columns are 1-indexed
          qikinkOrderSheet.getRange(i + 1, courierColumnIndex + 1).setValue(newCourierValue); // +1 because sheet rows/columns are 1-indexed
          Logger.log(orderNo + "(" + (i + 1) + "," + (trackingColumnIndex + 1) + ")" + "::" + oldTrackingValue + "::" + newTrackingValue)
          // 4.3.2 For trackings updates, add entry in Order Sync Logs sheet
          // addOrderSyncLog("Qikink", orderNo, "TrackingUpdate", foundQikinkOrder[0].oldTrackingValue, foundQikinkOrder[0].newTrackingValue, logNotes)
        }
      }
    }  

    //4.3.3 Get names from all the orders
    const foundQikinkOrder = qikinkOrders.filter(order => order.number === orderNo)  
    if (foundQikinkOrder && foundQikinkOrder.length > 0) {
      const newCustomerNameValue = foundQikinkOrder[0].shipping.first_name+" "+foundQikinkOrder[0].shipping.last_name
      if(newCustomerNameValue!=null) {
        qikinkOrderSheet.getRange(i + 1, customerNameColumnIndex + 1).setValue(newCustomerNameValue); // +1 because sheet rows/columns are 1-indexed
      }
    }
  }

  // 4.4 Get all order IDs from Qikink order sheet and log entry in Order Sync Logs sheet if the status is Exception/RTO Initiated/Returned
  for (let i = 1; i < data.length; i++) {
    const oldStatus = data[i][statusIndex].toString().trim().toLowerCase();
    const orderNo = data[i][5].toString().trim().toLowerCase();
    if (orderNo != "" && (oldStatus === "" || oldStatus === "prepaid" || oldStatus === "partially paid" || oldStatus === "cancelled")) {
      Logger.log(orderNo + "(" + (i + 1) + "," + (statusIndex + 1) + ")" + "::" + oldStatus)
      // 4.2 Check for each orders if any of it is showing Delivered in Qikink.
      const foundQikinkOrder = qikinkOrders.filter(order => order.number === orderNo)
      if (foundQikinkOrder && foundQikinkOrder.length > 0) {
        if (foundQikinkOrder[0].status === "Exception" || foundQikinkOrder[0].status === "RTO Initiated" || foundQikinkOrder[0].status === "Returned") {
          // this means we have found a Exception/RTO orders          
          Logger.log(orderNo + "(" + (i + 1) + "," + (statusIndex + 1) + ")" + "::" + foundQikinkOrder[0].status)
          Logger.log(foundQikinkOrder[0]);
          var logNotes = foundQikinkOrder[0].shipping.tracking_link;
          addOrderSyncLog("Qikink", orderNo, "ExceptionRTOCheck", foundQikinkOrder[0].status, "", logNotes)
        }
      }
    }
  }
}

//NOT IN USE
function processPendingOrders() {
  // 3.1 From the Qikink orders sheet, fetch the last order number
  var lastOrderPresent = fetchLastOrderFromQikinkSheet();
  Logger.log("Last Order = " + lastOrderPresent)
  if (lastOrderPresent != null) {
    var latestOrders = [];
    //Fetch Qikink Orders for 10 Pages (1 page = 10 orders)
    for (let pgno = 1; pgno <= 10; pgno++) {
      latestOrders.push(...fetchOrders(pgno));
    }
    Logger.log("Total Orders Fetched = " + latestOrders.length)

    //3.2 See if there are any new orders using Qikink APIs
    const filteredNewOrders = latestOrders.filter(order => {
      const orderPart = parseInt(order.number.replace("172086_", ""), 10);
      return orderPart > parseInt(lastOrderPresent.replace("172086_", ""), 10);
    });
    Logger.log("Total Filtered New Orders = " + filteredNewOrders.length)

    // 3.3 Insert new order rows in Qikink order sheet
    const qikinkOrderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Qikink Orders');
    var order = filteredNewOrders[1];
    Logger.log(order)
    //SrNo	Invoice #	Client	Order Date	Live Date	Order No	ORN	SKU	Product Price	Designs	Print Size	Print Price	Total	Handling	Shipping	IGST	CGST	SGST	Saved Value	Design	Total Cost	Sale Price	Amt Received	Amt Pending	Profit	Profit %	Status
    // qikinkOrderSheet.appendRow(['', 'Name', 'Financial Status', 'Update to Paid']);
    // orders.forEach(order => {
    //   qikinkOrderSheet.appendRow([order.id, order.name, order.financial_status, '']);
    // });

    //Traversing all Orders
    /*
    for (let index = latestOrders.length - 1; index >= 0; index--) {
      const order = latestOrders[index];
      Logger.log(`\n--- Order #${index + 1} ---`);
      Logger.log(`Order Number: ${order.number}`);
      Logger.log(`Status: ${order.status}`);
      Logger.log(`Payment Type: ${order.payment_type}`);
      Logger.log(`Total Order Value: ${order.total_order_value}`);      

      const shipping = order.shipping;
      if (shipping) {
        Logger.log(`Customer Name: ${shipping.first_name} ${shipping.last_name}`);
        Logger.log(`Email: ${shipping.email}`);
        Logger.log(`Phone: ${shipping.phone}`);
        Logger.log(`City: ${shipping.city}, Province: ${shipping.province}`);
      }

      const items = order.line_items || [];
      items.forEach((item, iIndex) => {
        Logger.log(`  Line Item #${iIndex + 1} - Quantity: ${item.quantity}, Price: ${item.price}`);
        const designs = item.designs || [];
        designs.forEach((design, dIndex) => {
          Logger.log(`    Design #${dIndex + 1}: ${design.design_code}`);
          Logger.log(`    URL: ${design.design_url}`);
          Logger.log(`    Printing Cost: ${design.printing_cost}`);
        });
      });
    };
    */
  }
}

//fetch the last order number present in Qikink Sheet
function fetchLastOrderFromQikinkSheet() {
  var sheetName = "Qikink Orders"
  var columnLetter = "F"
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  const column = sheet.getRange(columnLetter + ":" + columnLetter).getValues();

  // Loop from the bottom up to find the last non-empty cell
  for (let i = column.length - 1; i >= 0; i--) {
    if (column[i][0] !== "") {
      Logger.log("Last value in column " + columnLetter + ": " + column[i][0]);
      return column[i][0];  // Return the value
    }
  }
  Logger.log("Column " + columnLetter + " is empty.");
  return null;
}

function addOrderSyncLog(syncLane, orderId, changeType, oldState, newState, notes) {
  const syncLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Order Sync Logs');
  const newRow = [syncLane, new Date(), orderId, changeType, oldState, newState, notes]
  syncLogSheet.appendRow(newRow)
  Logger.log("*** New Sync Order Log Appended")
  Logger.log(newRow)
}

function fetchOrdersFromShopify() {
  const shopName = getCellValue("Automation", "H5"); // Replace with your store name
  const accessToken = getCellValue("Automation", "H6"); // Replace with your Admin API access token
  const apiVersion = getCellValue("Automation", "H7"); // Change if using a different API version
  const url = `https://${shopName}.myshopify.com/admin/api/${apiVersion}/orders.json?status=any&financial_status=any&limit=10`;

  const options = {
    method: 'get',
    headers: {
      'X-Shopify-Access-Token': accessToken,
      'Content-Type': 'application/json'
    }
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  const orders = data.orders;
  orders.forEach(order => {
    Logger.log(order.id + "::" + order.name + "::" + order.financial_status);
  });

}

function getShopifyOrderIdByOrderName(orderName) {
  const shopName = getCellValue("Automation", "H5"); // Replace with your store name
  const accessToken = getCellValue("Automation", "H6"); // Replace with your Admin API access token
  const apiVersion = getCellValue("Automation", "H7"); // Change if using a different API version

  // Make sure to include the "#" when querying
  const url = `https://${shopName}.myshopify.com/admin/api/${apiVersion}/orders.json?status=any&name=${encodeURIComponent(orderName)}`;

  const options = {
    method: "get",
    headers: {
      "X-Shopify-Access-Token": accessToken
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    if (data.orders && data.orders.length > 0) {
      const orderId = data.orders[0].id;
      Logger.log("For "+orderName+", Order ID Found: " + orderId);
      return orderId;
    } else {
      Logger.log("Order not found.");
      return null;
    }

  } catch (e) {
    Logger.log("Error fetching order: " + e);
    return null;
  }
}


function convertQikinkOrderNumberToShopifyNumber(orderNumber) {
  return orderNumber.replace("172086_", "#KC")
}
function convertShopifyOrderNumberToQikinkNumber(orderNumber) {
  return orderNumber.replace("#KC", "172086_")
}

function testing() {
  Logger.log(updateOrderStatusInShopify("#KC4687", "success", null))
}

function updateOrderStatusInShopify(orderName, status, amount) {
  Logger.log("***** updateOrderStatusInShopify(" + orderName + ", " + status + ", " + amount + ")")  
  var orderId = getShopifyOrderIdByOrderName(orderName);
  var result = "";
  if (orderId != null) {
    const shopName = getCellValue("Automation", "H5"); // Replace with your store name
    const accessToken = getCellValue("Automation", "H6"); // Replace with your Admin API access token
    const apiVersion = getCellValue("Automation", "H7"); // Change if using a different API version

    const url = `https://${shopName}.myshopify.com/admin/api/${apiVersion}/orders/${orderId}/transactions.json`;

    const payload = JSON.stringify({
      transaction: {
        kind: "sale",
        status: "success", // mark as successful payment
        amount: amount // If null, Shopify uses the full outstanding amount
      }
      // transaction: {
      //   kind: "capture",
      //   amount: 899,        
      // }
    });

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'X-Shopify-Access-Token': accessToken
      },
      payload: payload,
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(url, options);

      result = JSON.parse(response.getContentText());
      if (response.getResponseCode() === 201) {
        result += 'Order#' + orderId + ' is Marked ' + (status === 'success' ? 'Paid' : '');
      } else {
        result = 'ERROR: ' + (JSON.stringify(result.errors) || response.getContentText());
      }
    } catch (e) {
      result += 'ERROR: ' + e.message;
    }
  }
  else {
    result = "ERROR: Unable to fetch OrderID from OrderName=" + orderName
  }
  return result;
}

function getShopifyProducts() {
  const shopName = getCellValue("Automation", "H5"); // Replace with your store name
  const accessToken = getCellValue("Automation", "H6"); // Replace with your Admin API access token
  const apiVersion = getCellValue("Automation", "H7"); // Change if using a different API version

  const url = `https://${shopName}.myshopify.com/admin/api/${apiVersion}/products.json`;

  const options = {
    method: 'get',
    headers: {
      'X-Shopify-Access-Token': accessToken,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  // const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // sheet.clear(); // Clear existing content
  // sheet.appendRow(['Product Title', 'Product ID']); // Header row

  const products = data.products;
  Logger.log(products)

  // products.forEach(product => {
  //   sheet.appendRow([product.title, product.id]);
  // });
}

/**
 * Fetch all unfulfilled orders from Shopify
 */
function getUnfulfilledShopifyOrders() {
  const shopName = getCellValue("Automation", "H5"); // Replace with your store name
  const accessToken = getCellValue("Automation", "H6"); // Replace with your Admin API access token
  const apiVersion = getCellValue("Automation", "H7"); // Change if using a different API version

  const url = `https://${shopName}.myshopify.com/admin/api/${apiVersion}/orders.json?fulfillment_status=unfulfilled&status=any&limit=3`;
  
  const options = {
    method: 'get',
    headers: {
      'X-Shopify-Access-Token': accessToken,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    if (data.orders && data.orders.length > 0) {
      data.orders.forEach(order => {
        Logger.log(`Order ID: ${order.id} | Order Name: ${order.name} | Email: ${order.email} | Total: ${order.total_price} `);
      });      
      return data.orders;
    } else {
      Logger.log('No unfulfilled orders found.');
      return [];
    }
  } catch (error) {
    Logger.log('Error fetching unfulfilled orders: ' + error.message);
    return [];
  }
}

/* PLAN
* 1. Get all unfulfilled and non-cancelled orders from Shopify
* 2. Looping through them, get the Order number in Qikink
* 3. Fetch the corresponding order from qikink
* 4. Check if order is shipped. If yes, then upddate the details in Shopify
*
*/
function updateFulfillmentStatusFromQikinkToShopify() {
  // 1. Get all unfulfilled and non-cancelled orders from Shopify
  const ordersToProcess = getUnfulfilledShopifyOrders();
  Logger.log(`Total Unfulfilled Orders found: ${ordersToProcess.length}`)

  // 2. Looping through them, get the corresponding Order in Qikink
  ordersToProcess.forEach(order => {
    const qikinkOrderNo = convertShopifyOrderNumberToQikinkNumber(order.name);
    Logger.log(order);

    // 3. Fetch the corresponding order from qikink
    // const qikinkOrder = fetchQikinkOrderByID(qikinkOrderNo);
    // Logger.log(`Qikink Order: ${qikinkOrder[total_order_value]}`)
    // const test = fetchOrders(1);
    // test.forEach(ord => {
    //   Logger.log("***********")
    //   Logger.log(ord)
    // })
  }); 
}
