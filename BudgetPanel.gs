/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Budget Entry Panel - BudgetPanel.gs
 * 
 * This module provides a user-friendly panel interface to add income and expenses
 * directly to the Budget sheet with proper formatting and validation.
 */

const BudgetPanel = {

  /**
   * Show the budget entry panel
   */
  showBudgetEntryPanel() {
    try {
      console.log('Opening Budget Entry Panel...');
      
      const html = this._generateBudgetPanelHTML();
      const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(900)
        .setHeight(750)
        .setTitle('💰 Add Income & Expenses');
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Add Income & Expenses');
      
    } catch (error) {
      console.error('Error showing budget entry panel:', error);
      SpreadsheetApp.getUi().alert('Error', 'Failed to load budget entry panel: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  },

  /**
   * Add new budget entry (income or expense)
   */
  addBudgetEntry(entryData) {
    try {
      console.log('Adding budget entry:', entryData);
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const budgetSheet = spreadsheet.getSheetByName(SHEET_NAMES.BUDGET);
      
      if (!budgetSheet) {
        throw new Error('Budget sheet not found');
      }
      
      // Parse the entry data
      const data = JSON.parse(entryData);
      
      // Generate reference number
      const referenceNumber = this._generateReferenceNumber(data.type, data.category);
      
      // Format amount based on type (negative for expenses)
      const formattedAmount = this._formatAmountByType(data.amount, data.type);
      
      // Create budget entry row data
      const budgetRow = [
        data.date,                    // Date
        data.type,                    // Type (Income/Expense)
        data.description,             // Description
        formattedAmount,              // Amount (formatted with proper sign)
        data.category,                // Category
        data.paymentMethod,           // Payment Method
        referenceNumber,              // Reference Number (auto-generated)
        data.tenantGuest || '',       // Tenant/Guest
        data.receipt || ''            // Receipt
      ];
      
      // Add the new entry to the budget sheet
      const lastRow = budgetSheet.getLastRow();
      const newRow = lastRow + 1;
      budgetSheet.getRange(newRow, 1, 1, budgetRow.length).setValues([budgetRow]);
      
      // Apply currency formatting to the Amount column (Column D)
      const amountCell = budgetSheet.getRange(newRow, 4);
      amountCell.setNumberFormat('$#,##0.00');
      
      console.log(`Budget entry added successfully: ${data.type} - ${formattedAmount}`);
      return `✅ ${data.type} of ${formattedAmount} has been added to the budget successfully!`;
      
    } catch (error) {
      console.error('Error adding budget entry:', error);
      throw new Error('Failed to add budget entry: ' + error.message);
    }
  },

  /**
   * Get recent tenants and guests for the dropdown
   * @private
   */
  _getRecentTenantsAndGuests() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const names = new Set();
      
      // Add current tenants
      const tenantSheet = spreadsheet.getSheetByName(SHEET_NAMES.TENANT);
      if (tenantSheet && tenantSheet.getLastRow() > 1) {
        const tenantData = tenantSheet.getDataRange().getValues();
        const headers = tenantData[0];
        const nameCol = headers.indexOf('Current Tenant Name');
        const statusCol = headers.indexOf('Room Status');
        
        for (let i = 1; i < tenantData.length; i++) {
          if (tenantData[i][statusCol] === 'Occupied' && tenantData[i][nameCol]) {
            names.add(tenantData[i][nameCol]);
          }
        }
      }
      
      // Add recent guests
      const guestSheet = spreadsheet.getSheetByName(SHEET_NAMES.GUEST_ROOMS);
      if (guestSheet && guestSheet.getLastRow() > 1) {
        const guestData = guestSheet.getDataRange().getValues();
        const headers = guestData[0];
        const nameCol = headers.indexOf('Current Guest');
        
        for (let i = 1; i < guestData.length; i++) {
          if (guestData[i][nameCol]) {
            names.add(guestData[i][nameCol]);
          }
        }
      }
      
      // Convert to array and sort
      return Array.from(names).sort();
    } catch (error) {
      console.error('Error getting tenants and guests:', error);
      return [];
    }
  },

  /**
   * Generate reference number for budget entries
   * @private
   */
  _generateReferenceNumber(type, category) {
    const prefix = type === 'Income' ? 'INC' : 'EXP';
    const categoryCode = category.substring(0, 3).toUpperCase();
    const timestamp = Date.now().toString().slice(-6);
    return `${prefix}-${categoryCode}-${timestamp}`;
  },

  /**
   * Format amount based on type (negative for expenses, positive for income)
   * @private
   */
  _formatAmountByType(amount, type) {
    // Remove any existing currency symbols and clean the string
    let cleanAmount = amount.toString().replace(/[$,\s]/g, '');
    
    // Parse as float
    const numericAmount = parseFloat(cleanAmount);
    
    if (isNaN(numericAmount)) {
      return amount; // Return original if can't parse
    }
    
    // Make sure amount is positive first
    const positiveAmount = Math.abs(numericAmount);
    
    // Apply sign based on type
    const finalAmount = type === 'Expense' ? -positiveAmount : positiveAmount;
    
    // Format as currency
    return `$${Math.abs(finalAmount).toFixed(2)}${finalAmount < 0 ? '' : ''}`;
  },

  /**
   * Generate HTML for the budget entry panel
   * @private
   */
  _generateBudgetPanelHTML() {
    const tenantsAndGuests = this._getRecentTenantsAndGuests();
    const tenantsJson = JSON.stringify(tenantsAndGuests);
    
    return `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background: #1c4587; color: white; padding: 15px; margin: -20px -20px 20px -20px; }
        .entry-selector {
            background: #f8f9fa;
            border: 1px solid #ddd;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 20px;
        }
        .entry-form { 
            border: 1px solid #ddd; 
            border-radius: 8px; 
            margin: 15px 0; 
            padding: 20px; 
            background: #fff; 
        }
        .form-section {
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 1px solid #e0e0e0;
        }
        .form-section:last-child {
            border-bottom: none;
            margin-bottom: 0;
        }
        .section-title {
            font-size: 16px;
            font-weight: bold;
            color: #1c4587;
            margin-bottom: 20px;
            padding-bottom: 8px;
            border-bottom: 2px solid #1c4587;
        }
        .form-group { margin: 20px 0; }
        .form-group label { display: block; font-weight: bold; margin-bottom: 10px; color: #333; }
        .form-group input, .form-group select, .form-group textarea { 
            width: 100%; 
            padding: 12px; 
            border: 1px solid #ccc; 
            border-radius: 4px; 
            font-size: 14px;
            margin-bottom: 8px;
        }
        .form-group textarea {
            resize: vertical;
            min-height: 80px;
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
            gap: 20px;
            margin-bottom: 20px;
        }
        .form-row-three {
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            gap: 15px;
            margin-bottom: 20px;
        }
        .btn { 
            background: #1c4587; 
            color: white; 
            padding: 12px 25px; 
            border: none; 
            border-radius: 4px; 
            cursor: pointer; 
            font-size: 14px;
            font-weight: bold;
        }
        .btn:hover { background: #174a7e; }
        .btn-income { background: #22803c; }
        .btn-income:hover { background: #1a6b30; }
        .btn-expense { background: #dc3545; }
        .btn-expense:hover { background: #c82333; }
        .btn-secondary { 
            background: #6c757d; 
            margin-left: 10px;
        }
        .btn-secondary:hover { background: #5a6268; }
        .status { padding: 10px; margin: 10px 0; border-radius: 4px; }
        .status.success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .status.error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .required { color: #dc3545; }
        .currency-input { position: relative; }
        .currency-symbol { position: absolute; left: 15px; top: 50%; transform: translateY(-50%); color: #666; }
        .currency-input input { padding-left: 25px; }
        .type-income { background-color: #e8f5e8; }
        .type-expense { background-color: #ffe6e6; }
        .quick-amounts {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 10px;
            margin-top: 10px;
        }
        .quick-amounts button {
            padding: 8px 12px;
            font-size: 12px;
            background: #f8f9fa;
            border: 1px solid #dee2e6;
            border-radius: 4px;
            cursor: pointer;
        }
        .quick-amounts button:hover {
            background: #e9ecef;
        }
    </style>
</head>
<body>
    <div class="header">
        <h2>💰 Add Income & Expenses</h2>
        <p>Record income and expenses for the White House property</p>
    </div>
    
    <div id="status" class="status" style="display: none;"></div>
    
    <div class="entry-selector">
        <h3>Select Entry Type</h3>
        <div style="text-align: center; margin: 20px 0;">
            <button class="btn btn-income" onclick="showForm('income')" style="margin-right: 20px;">💰 Add Income</button>
            <button class="btn btn-expense" onclick="showForm('expense')">📤 Add Expense</button>
        </div>
    </div>
    
    <div class="entry-form" id="entry-form" style="display: none;">
        <div id="form-header" class="section-title"></div>
        
        <form id="budgetForm">
            <div class="form-section">
                <div class="section-title">Entry Details</div>
                <div class="form-row">
                    <div class="form-group">
                        <label>Date <span class="required">*</span>:</label>
                        <input type="date" id="entry-date" required>
                    </div>
                    <div class="form-group">
                        <label>Amount <span class="required">*</span>:</label>
                        <div class="currency-input">
                            <span class="currency-symbol">$</span>
                            <input type="number" id="entry-amount" placeholder="0.00" step="0.01" min="0" required>
                        </div>
                        <div class="quick-amounts">
                            <button type="button" onclick="setAmount(50)">$50</button>
                            <button type="button" onclick="setAmount(100)">$100</button>
                            <button type="button" onclick="setAmount(500)">$500</button>
                            <button type="button" onclick="setAmount(1000)">$1000</button>
                        </div>
                    </div>
                </div>
                
                <div class="form-group">
                    <label>Description <span class="required">*</span>:</label>
                    <textarea id="entry-description" placeholder="Describe the income or expense..." required></textarea>
                    <small>Provide a clear description of what this entry is for</small>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>Category <span class="required">*</span>:</label>
                        <select id="entry-category" required>
                            <option value="">-- Select Category --</option>
                            <option value="Rent">Rent</option>
                            <option value="Guest Revenue">Guest Revenue</option>
                            <option value="Maintenance">Maintenance</option>
                            <option value="Utilities">Utilities</option>
                            <option value="Insurance">Insurance</option>
                            <option value="Supplies">Supplies</option>
                            <option value="Marketing">Marketing</option>
                            <option value="Other">Other</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Payment Method <span class="required">*</span>:</label>
                        <select id="entry-payment-method" required>
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
                </div>
            </div>
            
            <div class="form-section">
                <div class="section-title">Additional Information</div>
                <div class="form-row">
                    <div class="form-group">
                        <label>Related to Tenant/Guest:</label>
                        <select id="entry-tenant-guest">
                            <option value="">-- Not applicable --</option>
                            ${tenantsAndGuests.map(name => `<option value="${name}">${name}</option>`).join('')}
                        </select>
                        <small>Select if this entry is related to a specific tenant or guest</small>
                    </div>
                    <div class="form-group">
                        <label>Receipt/Reference:</label>
                        <input type="text" id="entry-receipt" placeholder="Receipt number, invoice, etc.">
                        <small>Any receipt number or reference for this transaction</small>
                    </div>
                </div>
            </div>
            
            <div style="margin-top: 30px; text-align: center;">
                <button type="button" class="btn" id="submit-btn" onclick="submitEntry()">💰 Add Entry</button>
                <button type="button" class="btn btn-secondary" onclick="resetForm()">🔄 Clear Form</button>
            </div>
        </form>
    </div>
    
    <script>
        const tenantsAndGuests = ${tenantsJson};
        let currentEntryType = '';
        
        // Set today's date as default
        document.addEventListener('DOMContentLoaded', function() {
            const today = new Date().toISOString().split('T')[0];
            document.getElementById('entry-date').value = today;
        });
        
        function showForm(type) {
            currentEntryType = type;
            const form = document.getElementById('entry-form');
            const header = document.getElementById('form-header');
            const submitBtn = document.getElementById('submit-btn');
            
            form.style.display = 'block';
            
            if (type === 'income') {
                header.textContent = '💰 Add Income Entry';
                header.style.color = '#22803c';
                header.style.borderBottomColor = '#22803c';
                form.className = 'entry-form type-income';
                submitBtn.className = 'btn btn-income';
                submitBtn.textContent = '💰 Add Income';
            } else {
                header.textContent = '📤 Add Expense Entry';
                header.style.color = '#dc3545';
                header.style.borderBottomColor = '#dc3545';
                form.className = 'entry-form type-expense';
                submitBtn.className = 'btn btn-expense';
                submitBtn.textContent = '📤 Add Expense';
            }
            
            resetForm();
        }
        
        function setAmount(amount) {
            document.getElementById('entry-amount').value = amount.toFixed(2);
        }
        
        function submitEntry() {
            // Validate required fields
            const date = document.getElementById('entry-date').value;
            const amount = document.getElementById('entry-amount').value;
            const description = document.getElementById('entry-description').value;
            const category = document.getElementById('entry-category').value;
            const paymentMethod = document.getElementById('entry-payment-method').value;
            
            if (!date || !amount || !description || !category || !paymentMethod) {
                showStatus('Please fill in all required fields (marked with *).', 'error');
                return;
            }
            
            // Validate amount is a number
            if (isNaN(parseFloat(amount)) || parseFloat(amount) <= 0) {
                showStatus('Please enter a valid amount.', 'error');
                return;
            }
            
            const entryData = {
                type: currentEntryType === 'income' ? 'Income' : 'Expense',
                date: date,
                amount: parseFloat(amount).toFixed(2),
                description: description,
                category: category,
                paymentMethod: paymentMethod,
                tenantGuest: document.getElementById('entry-tenant-guest').value,
                receipt: document.getElementById('entry-receipt').value
            };
            
            showStatus('Adding entry to budget...', 'success');
            
            google.script.run
                .withSuccessHandler(function(result) {
                    showStatus(result, 'success');
                    resetForm();
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, 'error');
                })
                .addBudgetEntry(JSON.stringify(entryData));
        }
        
        function resetForm() {
            document.getElementById('budgetForm').reset();
            document.getElementById('entry-date').value = new Date().toISOString().split('T')[0];
            showStatus('Form cleared. Ready for new entry.', 'success');
        }
        
        function showStatus(message, type) {
            const status = document.getElementById('status');
            status.textContent = message;
            status.className = 'status ' + type;
            status.style.display = 'block';
            
            setTimeout(() => {
                if (type === 'success' && (message.includes('added') || message.includes('cleared'))) {
                    status.style.display = 'none';
                }
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
function showBudgetEntryPanel() {
  return BudgetPanel.showBudgetEntryPanel();
}

/**
 * Server-side function called from the HTML panel
 */
function addBudgetEntry(entryData) {
  return BudgetPanel.addBudgetEntry(entryData);
}
