/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Payment Recording Panel - PaymentPanel.gs
 * 
 * This module provides a user-friendly panel interface to record tenant payments
 * and update the Tenant sheet with payment information.
 * 
 * IMPROVEMENTS:
 * - Format amounts as currency in budget sheet
 * - Generate receipt numbers based on tenant
 * - Align payment method dropdowns with budget sheet options
 */

const PaymentPanel = {

  /**
   * Show the payment recording panel
   */
  showPaymentRecordingPanel() {
    try {
      console.log('Opening Payment Recording Panel...');
      
      const tenants = this._getTenantsForPayment();
      
      if (tenants.length === 0) {
        SpreadsheetApp.getUi().alert(
          'No Tenants Found',
          'No occupied rooms found for payment recording. Make sure you have tenants assigned to rooms.',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        return;
      }
      
      const html = this._generatePaymentPanelHTML(tenants);
      const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(900)
        .setHeight(700)
        .setTitle('💰 Record Tenant Payments');
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Record Tenant Payments');
      
    } catch (error) {
      console.error('Error showing payment panel:', error);
      SpreadsheetApp.getUi().alert('Error', 'Failed to load payment panel: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  },

  /**
   * Record payment for a tenant
   */
  recordTenantPayment(paymentData) {
    try {
      console.log('Recording payment:', paymentData);
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const tenantSheet = spreadsheet.getSheetByName(SHEET_NAMES.TENANT);
      const budgetSheet = spreadsheet.getSheetByName(SHEET_NAMES.BUDGET);
      
      if (!tenantSheet) {
        throw new Error('Tenant sheet not found');
      }
      
      // Parse the payment data
      const data = JSON.parse(paymentData);
      
      // Find the tenant row
      const tenantRowIndex = this._findTenantRow(tenantSheet, data.roomNumber, data.tenantName);
      if (tenantRowIndex === -1) {
        throw new Error(`Tenant ${data.tenantName} in Room ${data.roomNumber} not found`);
      }
      
      // Update tenant payment information
      this._updateTenantPaymentInfo(tenantSheet, tenantRowIndex, data);
      
      // Add payment record to budget sheet
      if (budgetSheet) {
        this._addPaymentToBudget(budgetSheet, data);
      }
      
      console.log(`Payment recorded successfully for ${data.tenantName} - ${data.amount}`);
      return `✅ Payment of ${data.amount} recorded for ${data.tenantName} in Room ${data.roomNumber}`;
      
    } catch (error) {
      console.error('Error recording payment:', error);
      throw new Error('Failed to record payment: ' + error.message);
    }
  },

  /**
   * Get all tenants that can receive payments
   * @private
   */
  _getTenantsForPayment() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const tenantSheet = spreadsheet.getSheetByName(SHEET_NAMES.TENANT);
      
      if (!tenantSheet || tenantSheet.getLastRow() <= 1) {
        return [];
      }
      
      const data = tenantSheet.getDataRange().getValues();
      const headers = data[0];
      const tenants = [];
      
      // Find column indices
      const roomNumberCol = headers.indexOf('Room Number');
      const rentalPriceCol = headers.indexOf('Rental Price');
      const negotiatedPriceCol = headers.indexOf('Negotiated Price');
      const nameCol = headers.indexOf('Current Tenant Name');
      const emailCol = headers.indexOf('Tenant Email');
      const statusCol = headers.indexOf('Room Status');
      const lastPaymentCol = headers.indexOf('Last Payment Date');
      const paymentStatusCol = headers.indexOf('Payment Status');
      
      // Process each tenant (skip header row)
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // Only include occupied rooms with tenant names
        if (row[statusCol] && row[statusCol].toString().toLowerCase() === 'occupied' && row[nameCol]) {
          const rentAmount = row[negotiatedPriceCol] || row[rentalPriceCol] || 0;
          const lastPayment = row[lastPaymentCol] ? new Date(row[lastPaymentCol]).toLocaleDateString() : 'None';
          
          tenants.push({
            roomNumber: row[roomNumberCol],
            tenantName: row[nameCol],
            email: row[emailCol] || '',
            rentAmount: rentAmount,
            lastPaymentDate: lastPayment,
            paymentStatus: row[paymentStatusCol] || 'Unknown',
            rowIndex: i + 1 // Store 1-based row index
          });
        }
      }
      
      // Sort by room number
      tenants.sort((a, b) => a.roomNumber.toString().localeCompare(b.roomNumber.toString()));
      
      console.log(`Found ${tenants.length} tenants for payment recording`);
      return tenants;
      
    } catch (error) {
      console.error('Error getting tenants for payment:', error);
      return [];
    }
  },

  /**
   * Find tenant row by room number and name
   * @private
   */
  _findTenantRow(sheet, roomNumber, tenantName) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const roomNumberCol = headers.indexOf('Room Number');
    const nameCol = headers.indexOf('Current Tenant Name');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][roomNumberCol].toString() === roomNumber.toString() && 
          data[i][nameCol].toString().trim() === tenantName.toString().trim()) {
        return i + 1; // Return 1-based row index
      }
    }
    
    return -1; // Not found
  },

  /**
   * Update tenant payment information in the sheet
   * @private
   */
  _updateTenantPaymentInfo(sheet, rowIndex, paymentData) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const lastPaymentCol = headers.indexOf('Last Payment Date');
    const paymentStatusCol = headers.indexOf('Payment Status');
    
    // Update last payment date
    if (lastPaymentCol !== -1) {
      sheet.getRange(rowIndex, lastPaymentCol + 1).setValue(paymentData.paymentDate);
    }
    
    // Update payment status
    let newStatus = 'Current'; // Default status
    
    if (paymentStatusCol !== -1) {
      // Determine status based on payment type
      if (paymentData.paymentType === 'Partial Payment') {
        newStatus = 'Partial';
      } else if (paymentData.paymentType === 'Late Payment') {
        newStatus = 'Current'; // Late payment brings them current
      } else {
        newStatus = 'Current';
      }
      
      sheet.getRange(rowIndex, paymentStatusCol + 1).setValue(newStatus);
    }
    
    console.log(`Updated payment info for row ${rowIndex}: Date=${paymentData.paymentDate}, Status=${newStatus}`);
  },

  /**
   * Generate reference number for payment tracking
   * @private
   */
  _generateReferenceNumber(tenantName, roomNumber, paymentDate) {
    // Create a reference like: "PAY-EDUARDO-R101-0808"
    const cleanName = tenantName.split(' ')[0].toUpperCase(); // First name only
    const cleanRoom = roomNumber.toString().padStart(3, '0'); // Ensure 3 digits
    const dateStr = new Date(paymentDate).toLocaleDateString('en-US', {
      month: '2-digit',
      day: '2-digit'
    }).replace('/', ''); // MMDD format
    
    return `PAY-${cleanName}-R${cleanRoom}-${dateStr}`;
  },

  /**
   * Generate receipt number based on tenant and date
   * @private
   */
  _generateReceiptNumber(tenantName, roomNumber, paymentDate) {
    // Create a receipt number like: "REC-EDUARDO-R101-240808-001"
    const cleanName = tenantName.split(' ')[0].toUpperCase(); // First name only
    const cleanRoom = roomNumber.toString().padStart(3, '0'); // Ensure 3 digits
    const date = new Date(paymentDate);
    const dateStr = date.getFullYear().toString().slice(-2) + // YY
                   (date.getMonth() + 1).toString().padStart(2, '0') + // MM
                   date.getDate().toString().padStart(2, '0'); // DD
    
    // Add a sequential number (could be enhanced to check existing receipts)
    const sequentialNumber = '001'; // For now, static. Could be improved to be dynamic
    
    return `REC-${cleanName}-R${cleanRoom}-${dateStr}-${sequentialNumber}`;
  },

  /**
   * Format amount as currency string
   * @private
   */
  _formatAsCurrency(amount) {
    // Remove any existing currency symbols and clean the string
    let cleanAmount = amount.toString().replace(/[$,\s]/g, '');
    
    // Parse as float
    const numericAmount = parseFloat(cleanAmount);
    
    if (isNaN(numericAmount)) {
      return amount; // Return original if can't parse
    }
    
    // Format as currency with $ symbol
    return `$${numericAmount.toFixed(2)}`;
  },

  /**
   * Add payment record to budget sheet with improved formatting
   * @private
   */
  _addPaymentToBudget(budgetSheet, paymentData) {
    try {
      // Generate reference number and receipt number automatically
      const referenceNumber = this._generateReferenceNumber(
        paymentData.tenantName, 
        paymentData.roomNumber, 
        paymentData.paymentDate
      );
      
      const receiptNumber = this._generateReceiptNumber(
        paymentData.tenantName,
        paymentData.roomNumber,
        paymentData.paymentDate
      );
      
      // Format amount as currency
      const formattedAmount = this._formatAsCurrency(paymentData.amount);
      
      // Create budget entry for the payment
      const budgetRow = [
        paymentData.paymentDate,           // Date
        'Income',                          // Type
        `Rent Payment - Room ${paymentData.roomNumber} - ${paymentData.tenantName}`, // Description
        formattedAmount,                   // Amount (formatted as currency)
        'Rent',                           // Category
        paymentData.paymentMethod,        // Payment Method
        referenceNumber,                  // Reference Number (auto-generated)
        paymentData.tenantName,           // Tenant/Guest
        receiptNumber                     // Receipt (auto-generated)
      ];
      
      // Add to the end of the budget sheet
      const lastRow = budgetSheet.getLastRow();
      const newRow = lastRow + 1;
      budgetSheet.getRange(newRow, 1, 1, budgetRow.length).setValues([budgetRow]);
      
      // Apply currency formatting to the Amount column (Column D)
      const amountCell = budgetSheet.getRange(newRow, 4);
      amountCell.setNumberFormat('$#,##0.00');
      
      console.log(`Added payment to budget sheet: ${formattedAmount} from ${paymentData.tenantName} (Ref: ${referenceNumber}, Receipt: ${receiptNumber})`);
      
    } catch (error) {
      console.error('Error adding payment to budget:', error);
      // Don't throw error here - payment recording should still succeed
    }
  },

  /**
   * Generate HTML for the payment recording panel with aligned dropdown options
   * @private
   */
  _generatePaymentPanelHTML(tenants) {
    const tenantsJson = JSON.stringify(tenants);
    
    return `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background: #22803c; color: white; padding: 15px; margin: -20px -20px 20px -20px; }
        .tenant-selector {
            background: #f8f9fa;
            border: 1px solid #ddd;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 20px;
        }
        .tenant-card { 
            border: 1px solid #ddd; 
            border-radius: 8px; 
            margin: 15px 0; 
            padding: 15px; 
            background: #f9f9f9; 
            display: none;
        }
        .tenant-header { 
            background: #e8f5e8; 
            padding: 10px; 
            margin: -15px -15px 15px -15px; 
            border-radius: 8px 8px 0 0; 
            font-weight: bold; 
        }
        .tenant-info { 
            display: grid; 
            grid-template-columns: 1fr 1fr; 
            gap: 15px; 
            margin: 15px 0;
            padding: 15px;
            background: #ffffff;
            border-radius: 5px;
            border: 1px solid #e0e0e0;
        }
        .info-item { margin: 5px 0; }
        .info-item strong { color: #22803c; }
        .payment-section { 
            background: #fff; 
            border: 1px solid #ccc; 
            border-radius: 5px; 
            padding: 20px; 
            margin-top: 15px; 
        }
        .form-group { margin: 20px 0; }
        .form-group label { display: block; font-weight: bold; margin-bottom: 10px; color: #333; }
        .form-group input, .form-group select { 
            width: calc(100% - 24px); 
            padding: 12px; 
            border: 1px solid #ccc; 
            border-radius: 4px; 
            font-size: 14px;
            margin-bottom: 8px;
            box-sizing: border-box;
        }
        .form-group small {
            color: #666; 
            font-size: 11px; 
            margin-top: 8px; 
            display: block;
            line-height: 1.4;
        }
        .form-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 25px;
            margin-bottom: 25px;
        }
        .form-row-three {
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            gap: 15px;
            margin-bottom: 20px;
        }
        .btn { 
            background: #22803c; 
            color: white; 
            padding: 12px 25px; 
            border: none; 
            border-radius: 4px; 
            cursor: pointer; 
            font-size: 14px;
            font-weight: bold;
        }
        .btn:hover { background: #1a6b30; }
        .btn-secondary { 
            background: #6c757d; 
            margin-left: 10px;
        }
        .btn-secondary:hover { background: #5a6268; }
        .status { padding: 10px; margin: 10px 0; border-radius: 4px; }
        .status.success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .status.error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .no-tenants { text-align: center; color: #666; margin: 50px 0; }
        .instruction { color: #666; font-style: italic; margin-bottom: 20px; text-align: center; }
        .payment-status { padding: 5px 10px; border-radius: 3px; font-size: 12px; font-weight: bold; }
        .status-current { background: #d4edda; color: #155724; }
        .status-late { background: #fff3cd; color: #856404; }
        .status-overdue { background: #f8d7da; color: #721c24; }
        .status-partial { background: #cce5ff; color: #004085; }
        .amount-display { font-size: 18px; font-weight: bold; color: #22803c; }
        .currency-input { 
            position: relative; 
            width: 100%;
        }
        .currency-symbol { 
            position: absolute; 
            left: 15px; 
            top: 50%; 
            transform: translateY(-50%); 
            color: #666; 
            z-index: 1;
            pointer-events: none;
        }
        .currency-input input { 
            padding-left: 35px;
            width: calc(100% - 24px);
        }
    </style>
</head>
<body>
    <div class="header">
        <h2>💰 Record Tenant Payments</h2>
        <p>Select a tenant to record their rent payment</p>
    </div>
    
    <div id="status" class="status" style="display: none;"></div>
    
    ${tenants.length === 0 ? `
        <div class="no-tenants">
            <h3>No Tenants Found</h3>
            <p>No occupied rooms with tenants found for payment recording.</p>
        </div>
    ` : `
        <div class="tenant-selector">
            <h3>Select Tenant for Payment</h3>
            <div class="instruction">Choose a tenant from the dropdown below to record their payment.</div>
            <div class="form-group">
                <label for="tenant-selector">Tenant:</label>
                <select id="tenant-selector" onchange="showSelectedTenant()">
                    <option value="">-- Choose a tenant --</option>
                    ${tenants.map((tenant, index) => `
                        <option value="${index}">Room ${tenant.roomNumber} - ${tenant.tenantName} (${tenant.paymentStatus})</option>
                    `).join('')}
                </select>
            </div>
        </div>
        
        ${tenants.map((tenant, index) => `
            <div class="tenant-card" id="tenant-${index}">
                <div class="tenant-header">
                    Payment for ${tenant.tenantName} - Room ${tenant.roomNumber}
                </div>
                
                <div class="tenant-info">
                    <div>
                        <div class="info-item"><strong>Room Number:</strong> ${tenant.roomNumber}</div>
                        <div class="info-item"><strong>Tenant Name:</strong> ${tenant.tenantName}</div>
                        <div class="info-item"><strong>Email:</strong> ${tenant.email || 'Not provided'}</div>
                    </div>
                    <div>
                        <div class="info-item"><strong>Rent Amount:</strong> <span class="amount-display">${tenant.rentAmount}</span></div>
                        <div class="info-item"><strong>Last Payment:</strong> ${tenant.lastPaymentDate}</div>
                        <div class="info-item"><strong>Status:</strong> <span class="payment-status status-${tenant.paymentStatus.toLowerCase()}">${tenant.paymentStatus}</span></div>
                    </div>
                </div>
                
                <div class="payment-section">
                    <h4 style="margin-bottom: 25px; color: #22803c;">💳 Payment Details</h4>
                    
                    <div class="form-row">
                        <div class="form-group">
                            <label>Payment Amount <span style="color: #dc3545;">*</span>:</label>
                            <div class="currency-input">
                                <span class="currency-symbol">$</span>
                                <input type="number" id="amount-${index}" placeholder="1200.00" step="0.01" min="0" required>
                            </div>
                            <small>Enter the payment amount (dollars and cents)</small>
                        </div>
                        <div class="form-group">
                            <label>Payment Date <span style="color: #dc3545;">*</span>:</label>
                            <input type="date" id="date-${index}" required>
                        </div>
                    </div>
                    
                    <div class="form-row">
                        <div class="form-group">
                            <label>Payment Method <span style="color: #dc3545;">*</span>:</label>
                            <select id="method-${index}" required>
                                <option value="">-- Select Method --</option>
                                <option value="Cash">Cash</option>
                                <option value="Bank Transfer">Bank Transfer</option>
                                <option value="Credit Card">Credit Card</option>
                                <option value="Debit Card">Debit Card</option>
                                <option value="Check">Check</option>
                                <option value="PayPal">PayPal</option>
                                <option value="Venmo">Venmo</option>
                                <option value="Zelle">Zelle</option>
                            </select>
                        </div>
                        <div class="form-group">
                            <label>Payment Type:</label>
                            <select id="type-${index}">
                                <option value="Regular Payment">Regular Payment</option>
                                <option value="Late Payment">Late Payment</option>
                                <option value="Partial Payment">Partial Payment</option>
                                <option value="Security Deposit">Security Deposit</option>
                                <option value="Other">Other</option>
                            </select>
                        </div>
                    </div>
                    
                    <div class="form-row">
                        <div class="form-group">
                            <label>Quick Amount:</label>
                            <button type="button" class="btn btn-secondary" onclick="setFullRent(${index})" style="width: 100%; margin: 0;">Full Rent (${tenant.rentAmount})</button>
                        </div>
                        <div class="form-group">
                            <label>Notes:</label>
                            <input type="text" id="notes-${index}" placeholder="Optional payment notes or comments">
                        </div>
                    </div>
                    
                    <div style="margin-top: 30px; text-align: center;">
                        <button class="btn" onclick="recordPayment(${index})">💰 Record Payment</button>
                        <button class="btn btn-secondary" onclick="clearForm(${index})">🔄 Clear Form</button>
                    </div>
                </div>
            </div>
        `).join('')}
    `}
    
    <script>
        const tenants = ${tenantsJson};
        
        // Set today's date as default
        document.addEventListener('DOMContentLoaded', function() {
            const today = new Date().toISOString().split('T')[0];
            tenants.forEach((tenant, index) => {
                const dateField = document.getElementById('date-' + index);
                if (dateField) {
                    dateField.value = today;
                }
            });
        });
        
        function showSelectedTenant() {
            const selectedIndex = document.getElementById('tenant-selector').value;
            
            // Hide all tenant cards
            tenants.forEach((tenant, index) => {
                document.getElementById('tenant-' + index).style.display = 'none';
            });
            
            // Show selected tenant card
            if (selectedIndex !== '') {
                document.getElementById('tenant-' + selectedIndex).style.display = 'block';
            }
        }
        
        function setFullRent(index) {
            const tenant = tenants[index];
            // Extract numeric value from rent amount
            const rentValue = tenant.rentAmount.toString().replace(/[$,]/g, '');
            document.getElementById('amount-' + index).value = rentValue;
        }
        
        function recordPayment(index) {
            const tenant = tenants[index];
            const amount = document.getElementById('amount-' + index).value;
            const paymentDate = document.getElementById('date-' + index).value;
            const paymentMethod = document.getElementById('method-' + index).value;
            const paymentType = document.getElementById('type-' + index).value;
            const notes = document.getElementById('notes-' + index).value;
            
            // Validation
            if (!amount || !paymentDate || !paymentMethod) {
                showStatus('Please fill in all required fields (Amount, Date, and Payment Method).', 'error');
                return;
            }
            
            // Validate amount is a number
            if (isNaN(parseFloat(amount)) || parseFloat(amount) <= 0) {
                showStatus('Please enter a valid payment amount.', 'error');
                return;
            }
            
            // Format amount as currency for display
            const formattedAmount = '$' + parseFloat(amount).toFixed(2);
            
            const paymentData = {
                roomNumber: tenant.roomNumber,
                tenantName: tenant.tenantName,
                amount: formattedAmount, // Send formatted amount
                paymentDate: paymentDate,
                paymentMethod: paymentMethod,
                paymentType: paymentType,
                notes: notes
            };
            
            showStatus('Recording payment...', 'success');
            
            google.script.run
                .withSuccessHandler(function(result) {
                    showStatus(result, 'success');
                    clearForm(index);
                    // Update the tenant selector to reflect payment
                    updateTenantSelector(index);
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, 'error');
                })
                .recordTenantPayment(JSON.stringify(paymentData));
        }
        
        function clearForm(index) {
            document.getElementById('amount-' + index).value = '';
            document.getElementById('date-' + index).value = new Date().toISOString().split('T')[0];
            document.getElementById('method-' + index).value = '';
            document.getElementById('type-' + index).value = 'Regular Payment';
            document.getElementById('notes-' + index).value = '';
        }
        
        function updateTenantSelector(index) {
            // Update the dropdown text to show payment was recorded
            const option = document.querySelector('#tenant-selector option[value="' + index + '"]');
            if (option) {
                const tenant = tenants[index];
                option.textContent = 'Room ' + tenant.roomNumber + ' - ' + tenant.tenantName + ' (Payment Recorded ✅)';
            }
        }
        
        function showStatus(message, type) {
            const status = document.getElementById('status');
            status.textContent = message;
            status.className = 'status ' + type;
            status.style.display = 'block';
            
            setTimeout(() => {
                status.style.display = 'none';
            }, 5000);
        }
    </script>
</body>
</html>
    `;
  }
};

/**
 * Wrapper function for menu integration
 */
function showPaymentRecordingPanel() {
  return PaymentPanel.showPaymentRecordingPanel();
}

/**
 * Server-side function called from the HTML panel
 */
function recordTenantPayment(paymentData) {
  return PaymentPanel.recordTenantPayment(paymentData);
}
