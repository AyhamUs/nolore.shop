/**
 * NOLORE Store - Google Apps Script
 * 
 * SETUP INSTRUCTIONS:
 * 
 * 1. Create a new Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Delete any existing code and paste this entire script
 * 4. Click "Deploy" > "New deployment"
 * 5. Select type: "Web app"
 * 6. Set "Execute as": "Me"
 * 7. Set "Who has access": "Anyone"
 * 8. Click "Deploy"
 * 9. Copy the Web App URL and paste it in your index.html file
 *    (Replace 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE')
 * 
 * The script will automatically create headers in your sheet on the first order.
 */

// Configuration - Change this to your sheet name if different
const SHEET_NAME = 'Orders';

/**
 * Handle POST requests from the website
 */
function doPost(e) {
  try {
    const sheet = getOrCreateSheet();
    const data = JSON.parse(e.postData.contents);
    
    // Add order to sheet (this also sends the confirmation email)
    addOrder(sheet, data);
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, message: 'Order received!' }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('Error processing order:', error);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle GET requests (for testing)
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ 
      status: 'NOLORE Store API is running!',
      message: 'Send POST requests with order data.'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Get existing sheet or create new one with headers
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    
    // Add headers
    const headers = [
      'Order ID',
      'Date & Time',
      'Customer Name',
      'Email',
      'Phone',
      'Shipping Address',
      'Items',
      'Total ($)',
      'Payment Method',
      'Status',
      'Notes',
      'Items (JSON)'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format headers
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#000000')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    
    // Set column widths
    sheet.setColumnWidth(1, 120);  // Order ID
    sheet.setColumnWidth(2, 160);  // Date & Time
    sheet.setColumnWidth(3, 150);  // Customer Name
    sheet.setColumnWidth(4, 200);  // Email
    sheet.setColumnWidth(5, 140);  // Phone
    sheet.setColumnWidth(6, 300);  // Shipping Address
    sheet.setColumnWidth(7, 350);  // Items (hoodie variants with sizes)
    sheet.setColumnWidth(8, 80);   // Total
    sheet.setColumnWidth(9, 120);  // Payment Method
    sheet.setColumnWidth(10, 100); // Status
    sheet.setColumnWidth(11, 200); // Notes
    sheet.setColumnWidth(12, 250); // Items JSON
    
    // Freeze header row
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * Add order to the sheet
 */
function addOrder(sheet, data) {
  // Generate order ID
  const orderId = 'NL-' + new Date().getTime().toString(36).toUpperCase();
  
  // Format date
  const date = new Date(data.timestamp);
  const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM dd, yyyy HH:mm:ss');
  
  // Format payment method for display
  const paymentDisplay = data.paymentMethod === 'revtrak' ? '💳 RevTrak' : '💵 Cash on Delivery';
  
  // Prepare row data
  const rowData = [
    orderId,
    formattedDate,
    data.name,
    data.email,
    data.phone,
    data.address,
    data.items,
    data.total,
    paymentDisplay,
    'Pending',
    '',  // Notes column for manual tracking
    data.itemsJson || ''
  ];
  
  // Append to sheet
  sheet.appendRow(rowData);
  
  // Get the last row and apply formatting
  const lastRow = sheet.getLastRow();
  
  // Alternate row colors
  if (lastRow % 2 === 0) {
    sheet.getRange(lastRow, 1, 1, rowData.length).setBackground('#f8f8f8');
  }
  
  // Color code the payment method column
  if (data.paymentMethod === 'revtrak') {
    sheet.getRange(lastRow, 9).setBackground('#e8f4fd').setFontColor('#1a73e8');
  } else {
    sheet.getRange(lastRow, 9).setBackground('#e8f5e9').setFontColor('#2e7d32');
  }
  
  // Color code the status (now column 10)
  sheet.getRange(lastRow, 10).setBackground('#fff3cd').setFontColor('#856404');
  
  // Send confirmation email with order ID
  sendConfirmationEmail(data, orderId);
  
  return orderId;
}

/**
 * Send beautiful HTML confirmation email to customer
 */
function sendConfirmationEmail(data, orderId) {
  const subject = `NOLORE - Order Confirmed #${orderId}`;
  
  // Parse items for nicer display
  const itemsList = data.items.split(', ').map(item => `
    <tr>
      <td style="padding: 16px 0; border-bottom: 1px solid #eee; font-size: 15px; color: #333;">
        ${item}
      </td>
    </tr>
  `).join('');
  
  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin: 0; padding: 0; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; background-color: #f5f5f5;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f5f5f5; padding: 40px 20px;">
    <tr>
      <td align="center">
        <table width="600" cellpadding="0" cellspacing="0" style="background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);">
          
          <!-- Header -->
          <tr>
            <td style="background-color: #0a0a0a; padding: 40px; text-align: center;">
              <h1 style="margin: 0; font-size: 28px; font-weight: 700; letter-spacing: 8px; color: #ffffff;">NOLORE</h1>
              <p style="margin: 8px 0 0 0; font-size: 11px; letter-spacing: 3px; color: #888888; text-transform: lowercase;">style, uncomplicated</p>
            </td>
          </tr>
          
          <!-- Success Icon & Message -->
          <tr>
            <td style="padding: 50px 40px 30px 40px; text-align: center;">
              <div style="width: 70px; height: 70px; background-color: #0a0a0a; border-radius: 50%; margin: 0 auto 25px auto; line-height: 70px;">
                <span style="color: #ffffff; font-size: 32px;">✓</span>
              </div>
              <h2 style="margin: 0 0 10px 0; font-size: 26px; font-weight: 600; color: #0a0a0a;">Order Confirmed!</h2>
              <p style="margin: 0; font-size: 15px; color: #666666;">Thank you for your purchase, ${data.name.split(' ')[0]}.</p>
            </td>
          </tr>
          
          <!-- Order Number Box -->
          <tr>
            <td style="padding: 0 40px 30px 40px;">
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f8f8f8; border-radius: 8px;">
                <tr>
                  <td style="padding: 25px; text-align: center;">
                    <p style="margin: 0 0 8px 0; font-size: 12px; letter-spacing: 2px; color: #888888; text-transform: uppercase;">Order Number</p>
                    <p style="margin: 0; font-size: 24px; font-weight: 700; letter-spacing: 2px; color: #0a0a0a;">${orderId}</p>
                  </td>
                </tr>
                <tr>
                  <td style="padding: 0 25px 25px 25px; text-align: center; border-top: 1px solid #eee;">
                    <p style="margin: 15px 0 0 0; font-size: 12px; letter-spacing: 2px; color: #888888; text-transform: uppercase;">Payment Method</p>
                    <p style="margin: 5px 0 0 0; font-size: 16px; font-weight: 600; color: #0a0a0a;">${data.paymentMethod === 'revtrak' ? '💳 RevTrak' : '💵 Cash on Delivery'}</p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          
          <!-- Order Details -->
          <tr>
            <td style="padding: 0 40px 30px 40px;">
              <h3 style="margin: 0 0 20px 0; font-size: 13px; font-weight: 600; letter-spacing: 2px; color: #888888; text-transform: uppercase; border-bottom: 2px solid #0a0a0a; padding-bottom: 10px;">Order Details</h3>
              <table width="100%" cellpadding="0" cellspacing="0">
                ${itemsList}
              </table>
            </td>
          </tr>
          
          <!-- Total -->
          <tr>
            <td style="padding: 0 40px 30px 40px;">
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #0a0a0a; border-radius: 8px;">
                <tr>
                  <td style="padding: 20px 25px;">
                    <table width="100%" cellpadding="0" cellspacing="0">
                      <tr>
                        <td style="font-size: 14px; color: #888888; text-transform: uppercase; letter-spacing: 1px;">Total</td>
                        <td align="right" style="font-size: 28px; font-weight: 700; color: #ffffff;">$${data.total}</td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          
          <!-- Shipping Address -->
          <tr>
            <td style="padding: 0 40px 40px 40px;">
              <h3 style="margin: 0 0 15px 0; font-size: 13px; font-weight: 600; letter-spacing: 2px; color: #888888; text-transform: uppercase;">Shipping To</h3>
              <p style="margin: 0; font-size: 15px; line-height: 1.6; color: #333333;">
                ${data.name}<br>
                ${data.address.replace(/\n/g, '<br>')}
              </p>
            </td>
          </tr>
          
          <!-- What's Next -->
          <tr>
            <td style="padding: 0 40px 40px 40px;">
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f8f8f8; border-radius: 8px;">
                <tr>
                  <td style="padding: 25px;">
                    <h3 style="margin: 0 0 15px 0; font-size: 13px; font-weight: 600; letter-spacing: 2px; color: #0a0a0a; text-transform: uppercase;">What's Next?</h3>
                    <p style="margin: 0; font-size: 14px; line-height: 1.7; color: #666666;">
                      We're preparing your order with care. You'll receive a shipping confirmation email with tracking information once your package is on its way.
                    </p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          
          <!-- Footer -->
          <tr>
            <td style="background-color: #0a0a0a; padding: 30px 40px; text-align: center;">
              <p style="margin: 0 0 10px 0; font-size: 16px; font-weight: 600; letter-spacing: 4px; color: #ffffff;">NOLORE</p>
              <p style="margin: 0 0 20px 0; font-size: 12px; color: #666666;">Style, Uncomplicated.</p>
              <p style="margin: 0; font-size: 11px; color: #444444;">
                Questions? Reply to this email or contact us at support@nolore.store
              </p>
            </td>
          </tr>
          
        </table>
        
        <!-- Bottom Text -->
        <p style="margin: 30px 0 0 0; font-size: 11px; color: #999999; text-align: center;">
          © 2025 NOLORE. All rights reserved.
        </p>
      </td>
    </tr>
  </table>
</body>
</html>
  `;
  
  // Plain text fallback
  const paymentMethodText = data.paymentMethod === 'revtrak' ? 'RevTrak' : 'Cash on Delivery';
  const plainBody = `
NOLORE - Order Confirmed!

Order Number: ${orderId}
Payment Method: ${paymentMethodText}

Hi ${data.name},

Thank you for your order from NOLORE!

ORDER DETAILS:
${data.items}

TOTAL: $${data.total}

SHIPPING TO:
${data.address}

We're preparing your order with care. You'll receive a shipping confirmation email with tracking information once your package is on its way.

Style, Uncomplicated.
- NOLORE Team

Questions? Contact us at support@nolore.store
  `;
  
  // Send HTML email
  MailApp.sendEmail({
    to: data.email,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody
  });
  
  console.log('Confirmation email sent to:', data.email);
}

/**
 * Update order status (can be called from sheet or script)
 */
function updateOrderStatus(orderId, newStatus) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      const statusCell = sheet.getRange(i + 1, 10);  // Status is now column 10
      statusCell.setValue(newStatus);
      
      // Color code by status
      switch(newStatus.toLowerCase()) {
        case 'pending':
          statusCell.setBackground('#fff3cd').setFontColor('#856404');
          break;
        case 'processing':
          statusCell.setBackground('#cce5ff').setFontColor('#004085');
          break;
        case 'shipped':
          statusCell.setBackground('#d4edda').setFontColor('#155724');
          break;
        case 'delivered':
          statusCell.setBackground('#28a745').setFontColor('#ffffff');
          break;
        case 'cancelled':
          statusCell.setBackground('#f8d7da').setFontColor('#721c24');
          break;
      }
      
      return true;
    }
  }
  
  return false;
}

/**
 * Get order statistics (can be called from sheet)
 */
function getOrderStats() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  
  let totalOrders = data.length - 1; // Exclude header
  let totalRevenue = 0;
  let pendingOrders = 0;
  
  for (let i = 1; i < data.length; i++) {
    totalRevenue += parseFloat(data[i][7]) || 0;
    if (data[i][8] === 'Pending') pendingOrders++;
  }
  
  return {
    totalOrders,
    totalRevenue,
    pendingOrders,
    averageOrderValue: totalOrders > 0 ? (totalRevenue / totalOrders).toFixed(2) : 0
  };
}

/**
 * Create a custom menu in Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('NOLORE')
    .addItem('View Order Stats', 'showOrderStats')
    .addItem('Mark Selected as Shipped', 'markAsShipped')
    .addItem('Mark Selected as Delivered', 'markAsDelivered')
    .addSeparator()
    .addItem('Setup Sheet', 'setupSheet')
    .addToUi();
}

/**
 * Show order statistics in a dialog
 */
function showOrderStats() {
  const stats = getOrderStats();
  const ui = SpreadsheetApp.getUi();
  
  ui.alert('NOLORE Order Statistics',
    `Total Orders: ${stats.totalOrders}\n` +
    `Total Revenue: $${stats.totalRevenue.toFixed(2)}\n` +
    `Pending Orders: ${stats.pendingOrders}\n` +
    `Average Order Value: $${stats.averageOrderValue}`,
    ui.ButtonSet.OK);
}

/**
 * Mark selected rows as shipped
 */
function markAsShipped() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    if (row > 1) { // Skip header
      const orderId = sheet.getRange(row, 1).getValue();
      if (orderId) {
        updateOrderStatus(orderId, 'Shipped');
      }
    }
  }
}

/**
 * Mark selected rows as delivered
 */
function markAsDelivered() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    if (row > 1) { // Skip header
      const orderId = sheet.getRange(row, 1).getValue();
      if (orderId) {
        updateOrderStatus(orderId, 'Delivered');
      }
    }
  }
}

/**
 * Manual setup function
 */
function setupSheet() {
  getOrCreateSheet();
  SpreadsheetApp.getUi().alert('Setup Complete', 'The Orders sheet has been created!', SpreadsheetApp.getUi().ButtonSet.OK);
}
