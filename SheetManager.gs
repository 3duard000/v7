/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Sheet Manager - SheetManager.gs
 * 
 * This module handles all sheet creation, formatting, column widths,
 * dropdowns, and comprehensive conditional formatting.
 */

const SheetManager = {

  /**
   * Create all required sheets with headers and formatting
   */
  createRequiredSheets() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create all main sheets (including Guest Rooms)
    const mainSheets = ['TENANT', 'BUDGET', 'MAINTENANCE', 'GUEST_ROOMS'];
    
    mainSheets.forEach(key => {
      const sheetName = SHEET_NAMES[key];
      let sheet = spreadsheet.getSheetByName(sheetName);
      
      if (!sheet) {
        console.log(`Creating sheet: ${sheetName}`);
        sheet = spreadsheet.insertSheet(sheetName);
      }
      
      // Set headers if the sheet is empty or has different headers
      const headers = HEADERS[key];
      if (!headers) {
        console.log(`No headers defined for ${key}, skipping...`);
        return;
      }
      
      const lastRow = sheet.getLastRow();
      if (lastRow === 0) {
        // Sheet is completely empty, add headers
        this._addHeaders(sheet, headers);
        this._formatSheet(sheet, key);
        console.log(`Sheet "${sheetName}" created/updated with headers and formatting`);
      } else {
        // Check if headers match
        this._updateHeadersIfNeeded(sheet, headers, key);
      }
    });
  },

  /**
   * Add headers to a sheet
   * @private
   */
  _addHeaders(sheet, headers) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row with stronger blue
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1c4587'); // Stronger blue
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
  },

  /**
   * Update headers if they don't match expected headers
   * @private
   */
  _updateHeadersIfNeeded(sheet, headers, key) {
    const existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const existingHeadersString = existingHeaders.join('');
    const expectedHeadersString = headers.join('');
    
    if (existingHeadersString !== expectedHeadersString) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      this._formatSheet(sheet, key);
      console.log(`Sheet "${sheet.getName()}" headers updated`);
    }
  },

  /**
   * Apply formatting to a sheet (column widths, freezing, dropdowns, conditional formatting)
   * @private
   */
  _formatSheet(sheet, sheetKey) {
    // Set column widths
    this._setColumnWidths(sheet, sheetKey);
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Add dropdowns and conditional formatting
    this._addDropdownsAndFormatting(sheet, sheetKey);
  },

  /**
   * Set column widths for a sheet
   * @private
   */
  _setColumnWidths(sheet, sheetKey) {
    const columnWidths = COLUMN_WIDTHS[sheetKey];
    if (!columnWidths) return;
    
    // Apply column widths
    for (let i = 0; i < columnWidths.length; i++) {
      sheet.setColumnWidth(i + 1, columnWidths[i]);
    }
    
    console.log(`Column widths set for ${sheet.getName()}`);
  },

  /**
   * Add dropdown menus and conditional formatting to a sheet
   * @private
   */
  _addDropdownsAndFormatting(sheet, sheetKey) {
    switch(sheetKey) {
      case 'TENANT':
        this._formatTenantSheet(sheet);
        break;
      case 'GUEST_ROOMS':
        this._formatGuestRoomsSheet(sheet);
        break;
      case 'BUDGET':
        this._formatBudgetSheet(sheet);
        break;
      case 'MAINTENANCE':
        this._formatMaintenanceSheet(sheet);
        break;
    }
  },

  /**
   * Format Tenant sheet with dropdowns and conditional formatting
   * @private
   */
  _formatTenantSheet(sheet) {
    // Room Status dropdown (Column I - index 9)
    const roomStatusRange = sheet.getRange('I2:I1000');
    this._addDataValidation(roomStatusRange, DROPDOWN_OPTIONS.TENANT.ROOM_STATUS);
    
    // Payment Status dropdown (Column K - index 11)
    const paymentStatusRange = sheet.getRange('K2:K1000');
    this._addDataValidation(paymentStatusRange, DROPDOWN_OPTIONS.TENANT.PAYMENT_STATUS);
    
    // Conditional formatting for Room Status (Column I)
    const roomStatusRules = [
      this._createConditionalRule('Occupied', COLORS.LIGHT_GREEN, roomStatusRange), // Light green
      this._createConditionalRule('Vacant', COLORS.LIGHT_YELLOW, roomStatusRange), // Light yellow
      this._createConditionalRule('Maintenance', COLORS.LIGHT_RED, roomStatusRange), // Light red
      this._createConditionalRule('Reserved', COLORS.LIGHT_BLUE, roomStatusRange) // Light blue
    ];
    
    // Conditional formatting for Payment Status (Column K)
    const paymentStatusRules = [
      this._createConditionalRule('Current', COLORS.LIGHT_GREEN, paymentStatusRange), // Light green
      this._createConditionalRule('Late', COLORS.LIGHT_ORANGE, paymentStatusRange), // Light orange
      this._createConditionalRule('Overdue', COLORS.LIGHT_RED, paymentStatusRange), // Light red
      this._createConditionalRule('Partial', COLORS.LIGHT_YELLOW, paymentStatusRange) // Light yellow
    ];
    
    // Combine all rules
    const allTenantRules = roomStatusRules.concat(paymentStatusRules);
    sheet.setConditionalFormatRules(allTenantRules);
    
    console.log('✅ Tenant sheet dropdowns and formatting added');
  },

  /**
   * Format Guest Rooms sheet with dropdowns and conditional formatting
   * @private
   */
  _formatGuestRoomsSheet(sheet) {
    // Room Type dropdown (Column D)
    this._addDataValidation(sheet.getRange('D2:D1000'), DROPDOWN_OPTIONS.GUEST_ROOMS.ROOM_TYPE);
    
    // Status dropdown (Column J)
    const statusRange = sheet.getRange('J2:J1000');
    this._addDataValidation(statusRange, DROPDOWN_OPTIONS.GUEST_ROOMS.STATUS);
    
    // Payment Status dropdown (Column V)
    const paymentStatusRange = sheet.getRange('V2:V1000');
    this._addDataValidation(paymentStatusRange, DROPDOWN_OPTIONS.GUEST_ROOMS.PAYMENT_STATUS);
    
    // Booking Status dropdown (Column W)
    const bookingStatusRange = sheet.getRange('W2:W1000');
    this._addDataValidation(bookingStatusRange, DROPDOWN_OPTIONS.GUEST_ROOMS.BOOKING_STATUS);
    
    // Source dropdown (Column T)
    this._addDataValidation(sheet.getRange('T2:T1000'), DROPDOWN_OPTIONS.GUEST_ROOMS.SOURCE);
    
    // Conditional formatting for Room Status (Column J)
    const roomStatusRules = [
      this._createConditionalRule('Available', COLORS.LIGHT_GREEN, statusRange), // Light green
      this._createConditionalRule('Occupied', COLORS.LIGHT_BLUE, statusRange), // Light blue
      this._createConditionalRule('Reserved', COLORS.LIGHT_ORANGE, statusRange), // Light orange
      this._createConditionalRule('Maintenance', COLORS.LIGHT_RED, statusRange), // Light red
      this._createConditionalRule('Cleaning', COLORS.LIGHT_YELLOW, statusRange) // Light yellow
    ];
    
    // Conditional formatting for Payment Status (Column V)
    const paymentStatusRules = [
      this._createConditionalRule('Paid', COLORS.LIGHT_GREEN, paymentStatusRange), // Light green
      this._createConditionalRule('Pending', COLORS.LIGHT_YELLOW, paymentStatusRange), // Light yellow
      this._createConditionalRule('Deposit Paid', COLORS.LIGHT_BLUE, paymentStatusRange), // Light blue
      this._createConditionalRule('Cancelled', COLORS.LIGHT_RED, paymentStatusRange), // Light red
      this._createConditionalRule('Refunded', COLORS.LIGHT_ORANGE, paymentStatusRange) // Light orange
    ];
    
    // Conditional formatting for Booking Status (Column W)
    const bookingStatusRules = [
      this._createConditionalRule('Confirmed', COLORS.LIGHT_GREEN, bookingStatusRange), // Light green
      this._createConditionalRule('Checked-In', COLORS.LIGHT_BLUE, bookingStatusRange), // Light blue
      this._createConditionalRule('Checked-Out', '#f3f3f3', bookingStatusRange), // Light gray
      this._createConditionalRule('Tentative', COLORS.LIGHT_YELLOW, bookingStatusRange), // Light yellow
      this._createConditionalRule('Cancelled', COLORS.LIGHT_RED, bookingStatusRange), // Light red
      this._createConditionalRule('No-Show', COLORS.LIGHT_RED, bookingStatusRange) // Light red
    ];
    
    // Combine all rules
    const allGuestRules = roomStatusRules.concat(paymentStatusRules).concat(bookingStatusRules);
    sheet.setConditionalFormatRules(allGuestRules);
    
    console.log('✅ Guest Rooms sheet dropdowns and formatting added');
  },

  /**
   * Format Budget sheet with dropdowns and conditional formatting
   * @private
   */
  _formatBudgetSheet(sheet) {
    // Type dropdown (Column B)
    const typeRange = sheet.getRange('B2:B1000');
    this._addDataValidation(typeRange, DROPDOWN_OPTIONS.BUDGET.TYPE);
    
    // Category dropdown (Column E)
    this._addDataValidation(sheet.getRange('E2:E1000'), DROPDOWN_OPTIONS.BUDGET.CATEGORY);
    
    // Payment Method dropdown (Column F)
    this._addDataValidation(sheet.getRange('F2:F1000'), DROPDOWN_OPTIONS.BUDGET.PAYMENT_METHOD);
    
    // Amount range (Column D) - for conditional formatting
    const amountRange = sheet.getRange('D2:D1000');
    
    // Custom conditional formatting based on Type column
    // Income - format Amount column with light green background
    const incomeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="Income"')
      .setBackground(COLORS.LIGHT_GREEN) // Light green
      .setFontColor('#000000') // Black text
      .setRanges([amountRange])
      .build();
    
    // Expense - format Amount column with light red background  
    const expenseRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="Expense"')
      .setBackground(COLORS.LIGHT_RED) // Light red
      .setFontColor('#000000') // Black text
      .setRanges([amountRange])
      .build();
    
    // Conditional formatting for Type column itself
    const typeRules = [
      this._createConditionalRule('Income', COLORS.LIGHT_GREEN, typeRange), // Light green
      this._createConditionalRule('Expense', COLORS.LIGHT_RED, typeRange) // Light red
    ];
    
    // Combine all rules
    const allBudgetRules = [incomeRule, expenseRule].concat(typeRules);
    sheet.setConditionalFormatRules(allBudgetRules);
    
    console.log('✅ Budget sheet dropdowns and formatting added');
  },

  /**
   * Format Maintenance sheet with dropdowns and conditional formatting
   * @private
   */
  _formatMaintenanceSheet(sheet) {
    // Issue Type dropdown (Column D)
    this._addDataValidation(sheet.getRange('D2:D1000'), DROPDOWN_OPTIONS.MAINTENANCE.ISSUE_TYPE);
    
    // Priority dropdown (Column E)
    const priorityRange = sheet.getRange('E2:E1000');
    this._addDataValidation(priorityRange, DROPDOWN_OPTIONS.MAINTENANCE.PRIORITY);
    
    // Status dropdown (Column J)
    const statusRange = sheet.getRange('J2:J1000');
    this._addDataValidation(statusRange, DROPDOWN_OPTIONS.MAINTENANCE.STATUS);
    
    // Conditional formatting for Priority (Column E)
    const priorityRules = [
      this._createConditionalRule('Low', COLORS.LIGHT_GREEN, priorityRange), // Light green
      this._createConditionalRule('Medium', COLORS.LIGHT_YELLOW, priorityRange), // Light yellow
      this._createConditionalRule('High', COLORS.LIGHT_ORANGE, priorityRange), // Light orange
      this._createConditionalRule('Emergency', COLORS.LIGHT_RED, priorityRange) // Light red
    ];
    
    // Conditional formatting for Status (Column J)
    const statusRules = [
      this._createConditionalRule('Completed', COLORS.LIGHT_GREEN, statusRange), // Light green
      this._createConditionalRule('In Progress', COLORS.LIGHT_BLUE, statusRange), // Light blue
      this._createConditionalRule('Pending', COLORS.LIGHT_YELLOW, statusRange), // Light yellow
      this._createConditionalRule('Cancelled', '#f3f3f3', statusRange), // Light gray
      this._createConditionalRule('On Hold', COLORS.LIGHT_ORANGE, statusRange) // Light orange
    ];
    
    // Combine all maintenance rules
    const allMaintenanceRules = priorityRules.concat(statusRules);
    sheet.setConditionalFormatRules(allMaintenanceRules);
    
    console.log('✅ Maintenance sheet dropdowns and formatting added');
  },

  /**
   * Add data validation (dropdown) to a range
   * @private
   */
  _addDataValidation(range, options) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options)
      .setAllowInvalid(false)
      .build();
    range.setDataValidation(rule);
  },

  /**
   * Create a conditional formatting rule with black text for better readability
   * @private
   */
  _createConditionalRule(textValue, backgroundColor, range) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(textValue)
      .setBackground(backgroundColor)
      .setFontColor('#000000') // Black text for better readability
      .setRanges([range])
      .build();
  }
};
