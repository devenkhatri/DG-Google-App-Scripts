function addTrackingsToGS() {
  SpreadsheetApp.getActiveSpreadsheet().toast("KC Tools -> Process Orders for Status Updates Started", "Info");
  const qikinkOrderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Qikink Orders');
  const data = qikinkOrderSheet.getDataRange().getValues(); // Get all data in the sheet
  const header = data[0];
  const trackingColumnIndex = header.indexOf("TrackingLink"); 
  const awbColumnIndex = header.indexOf("AwbNo"); 
  const courierColumnIndex = header.indexOf("CourierPartner"); 

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
    const oldTrackingValue = data[i][trackingColumnIndex].toString().trim().toLowerCase();
    const orderNo = data[i][5].toString().trim().toLowerCase();

    // 4.1 Get all order IDs from Qikink order sheet, where status is blank or Prepaid or Partially Paid
    if (orderNo != "" && oldTrackingValue !== "") {
      const foundQikinkOrder = qikinkOrders.filter(order => order.number === orderNo && order.shipping.awb != null)
      if (foundQikinkOrder && foundQikinkOrder.length > 0) {
        const newAwbValue = foundQikinkOrder[0].shipping.awb
        const newCourierValue = foundQikinkOrder[0].shipping.courier_provider_name
        // Logger.log(foundQikinkOrder[0]);
        qikinkOrderSheet.getRange(i + 1, awbColumnIndex + 1).setValue(newAwbValue); // +1 because sheet rows/columns are 1-indexed
        qikinkOrderSheet.getRange(i + 1, courierColumnIndex + 1).setValue(newCourierValue); // +1 because sheet rows/columns are 1-indexed
      }
    }    
  }
}
/**
 * Get fulfillment orders from the order
 */
function getFulfillmentId(orderId) {
  const shopName = getCellValue("Automation", "H5"); // Replace with your store name
  const accessToken = getCellValue("Automation", "H6"); // Replace with your Admin API access token
  const apiVersion = getCellValue("Automation", "H7"); // Change if using a different API version

  const fulfillmentOrdersUrl = `https://${shopName}.myshopify.com/admin/api/${apiVersion}/orders/${orderId}/fulfillment_orders.json`;

  const options = {
    'method': 'get',
    'headers': {
      'X-Shopify-Access-Token': accessToken,
      'Content-Type': 'application/json'
    },
    'muteHttpExceptions': true
  };

  const response = UrlFetchApp.fetch(fulfillmentOrdersUrl, options);
  const fulfillmentOrders = JSON.parse(response.getContentText()).fulfillment_orders;

  if (!fulfillmentOrders || fulfillmentOrders.length === 0) {
    Logger.log("No fulfillment orders found for order " + orderId);
    return;
  }
  Logger.log(fulfillmentOrders);
  const fulfillmentOrderId = fulfillmentOrders[0].id; // Assuming you want to fulfill the first one  
  return fulfillmentOrderId;
}

/**
 * Creates a new fulfillment for a Shopify order and adds tracking info.
 */
function createNewFulfillmentWithTracking(orderId, trackingNumber, trackingUrl, trackingCompany) {
  const shopName = getCellValue("Automation", "H5"); // Replace with your store name
  const accessToken = getCellValue("Automation", "H6"); // Replace with your Admin API access token
  const apiVersion = getCellValue("Automation", "H7"); // Change if using a different API version

  const fulfillmentId = getFulfillmentId(orderId)  

  // Make sure to include the "#" when querying
  const url = `https://${shopName}.myshopify.com/admin/api/${apiVersion}/orders/${orderId}/fulfillments/${fulfillmentId}/update_tracking.json`;
  Logger.log(`URL ${url}`)

  const payload = {
    fulfillment: {
      tracking_info : {
        company: trackingCompany,
        number: trackingNumber,
        url: trackingUrl,
      },
      notify_customer: true      
    }
  };
  Logger.log(JSON.stringify(payload))
  const options = {
    method: 'post',
    headers: {
      'X-Shopify-Access-Token': accessToken,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log('HTTP Status: ' + responseCode);
    Logger.log('Raw Response: ' + responseText);
    const result = JSON.parse(response.getContentText());

    // Try to parse JSON only if response is not empty
    if (responseText.trim()) {
      const result = JSON.parse(responseText);

      if (responseCode === 200) {
        Logger.log('Tracking info updated successfully.');
        Logger.log(result);
      } else {
        Logger.log('Shopify API returned an error:');
        Logger.log(JSON.stringify(result, null, 2));
      }
    } else {
      Logger.log('Empty response body received from Shopify.');
    }
  } catch (error) {
    Logger.log('Error updating tracking info: ' + error.message);
  }
}

function addTrackingsToShopify() {
  SpreadsheetApp.getActiveSpreadsheet().toast("KC Tools -> Process Orders for Status Updates Started", "Info");
  const qikinkOrderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Qikink Orders');
  const data = qikinkOrderSheet.getDataRange().getValues(); // Get all data in the sheet
  const header = data[0];
  const trackingColumnIndex = header.indexOf("TrackingLink"); // Find the "status" column index
  const awbColumnIndex = header.indexOf("AwbNo"); 
  const courierColumnIndex = header.indexOf("CourierPartner"); 
  const shopifyFulfillmentColumnIndex = header.indexOf("ShopifyFulfillment"); // Find the "status" column index
  if (trackingColumnIndex === -1) {
    Logger.log("No 'TrackingLink' column found.");
    return;
  }

  for (let i = 4670; i < data.length; i++) {
    const oldTrackingValue = data[i][trackingColumnIndex].toString().trim().toLowerCase();
    const oldAwbValue = data[i][awbColumnIndex].toString().trim().toLowerCase();
    const oldCourierValue = data[i][courierColumnIndex].toString().trim().toLowerCase();
    const oldShopifyFulfillmentValue = data[i][shopifyFulfillmentColumnIndex].toString().trim().toLowerCase();
    const orderNo = data[i][5].toString().trim().toLowerCase();

    // Create Shopify Fulfillment if tracking is available but it is yet not notified to shopify
    if (orderNo != "" && oldTrackingValue !== "" && oldShopifyFulfillmentValue === "") {
      var orderId = getShopifyOrderIdByOrderName(convertQikinkOrderNumberToShopifyNumber(orderNo));
      Logger.log(orderNo + "(" + orderId + "," + oldAwbValue+ ","+  oldTrackingValue + ","+ oldCourierValue)  
      createNewFulfillmentWithTracking(orderId, oldAwbValue, oldTrackingValue, oldCourierValue)
    }    
  }
}