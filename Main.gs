/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Main Entry Point - Main.gs
 * 
 * This is the main entry point for the White House property management system.
 * It coordinates all the different modules and provides the main setup function.
 * Updated with reorganized menu structure and Budget Entry Panel integration.
 */

/**
 * Create custom menu when spreadsheet opens
 * UPDATED: Reorganized menu items in logical workflow order + Budget Entry Panel
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🏠 White House Manager')
    .addItem('⚙️ Initialize System', 'setupTenantManagement')
    .addSeparator()
    .addSubMenu(ui.createMenu('📝 Application Management')
      .addItem('📝 Process Applications', 'showApplicationProcessingPanel')
      .addItem('🚪 Process Move-Outs', 'showMoveOutProcessingPanel'))
    .addSubMenu(ui.createMenu('🏨 Guest Management')
      .addItem('⚡ Check Availability', 'showAvailabilityChecker')
      .addItem('📝 Create New Booking', 'showNewBookingPanel')
      .addItem('🏨 Process Online Reservations', 'showOnlineReservationPanel')
      .addItem('✅ Process Check-In', 'showCheckInPanel')
      .addItem('🚪 Process Check-Out', 'showCheckOutPanel'))
    .addSubMenu(ui.createMenu('💰 Financial Management')
      .addItem('💰 Record Payments', 'showPaymentRecordingPanel')
      .addItem('💵 Add Income/Expenses', 'showBudgetEntryPanel'))
    .addSubMenu(ui.createMenu('🔧 Property Management')
      .addItem('🔧 Add Maintenance Request', 'showMaintenanceRequestPanel'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Reports & Dashboards')
      .addItem('🏠 Management Dashboard', 'createManagementDashboard')
      .addItem('💰 Financial Dashboard', 'createFinancialDashboard')
      .addSeparator()
      .addItem('🔄 Refresh All Dashboards', 'refreshAllDashboards'))
    .addSubMenu(ui.createMenu('📧 Email Management')
      .addItem('💸 Send Rent Reminders', 'sendRentReminders')
      .addItem('⚠️ Send Late Payment Alerts', 'sendLatePaymentAlerts')
      .addItem('📋 Send Monthly Invoices', 'sendMonthlyInvoices')
      .addSeparator()
      .addItem('📧 Send All Payment Alerts', 'sendAllPaymentAlerts')
      .addSeparator()
      .addItem('🔄 Update Payment Status', 'updateAllPaymentStatuses'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🛠️ System Tools')
      .addItem('🔄 Rename Form Sheets', 'performDelayedSheetRenaming'))
    .addToUi();
}

/**
 * Main setup function - run this first to initialize everything
 * Updated to include delayed sheet renaming
 */
function setupTenantManagement() {
  try {
    console.log('Setting up Tenant Management System...');
    
    // Create sheets with headers and formatting
    SheetManager.createRequiredSheets();
    
    // Add sample data to demonstrate the system
    DataManager.addSampleData();
    
    // Create and link Google Forms
    FormManager.createGoogleForms();
    
    // Set up automated email triggers
    TriggerManager.setupTriggers();
    
    // Create initial dashboards
    Dashboard.createManagementDashboard();
    Dashboard.createFinancialDashboard();
    
    // **NEW: Set up delayed sheet renaming after forms are created**
    console.log('Setting up delayed sheet renaming...');
    FormManager.setupDelayedSheetRenaming();
    
    console.log('Setup completed successfully!');
    return 'Tenant Management System setup completed successfully! ' +
           'Dashboards created and will auto-refresh 3x daily. ' +
           'Form response sheets will be automatically renamed in 3 minutes. ' +
           'Check the execution log for form URLs.';
  } catch (error) {
    console.error('Setup failed:', error);
    throw new Error('Setup failed: ' + error.message);
  }
}

/**
 * Manual function to trigger sheet renaming
 * This can be called manually from the menu if needed
 */
function performDelayedSheetRenaming() {
  try {
    console.log('Manually triggering sheet renaming...');
    return FormManager.performDelayedSheetRenaming();
  } catch (error) {
    console.error('Manual sheet renaming failed:', error);
    throw new Error('Sheet renaming failed: ' + error.message);
  }
}

/**
 * Alternative approach: Force sheet creation by submitting test responses
 * Run this manually if the delayed approach doesn't work
 */
function forceFormSheetCreation() {
  try {
    console.log('Forcing form sheet creation...');
    
    // This will submit dummy responses to all forms to force sheet creation
    FormManager._triggerFormResponseSheetCreation();
    
    // Wait for sheets to be created
    console.log('Waiting 30 seconds for sheets to be created...');
    Utilities.sleep(30000);
    
    // Then rename them
    FormManager._renameFormResponseSheets();
    
    return 'Form sheets created and renamed successfully!';
    
  } catch (error) {
    console.error('Force sheet creation failed:', error);
    throw new Error('Force sheet creation failed: ' + error.message);
  }
}

/**
 * Trigger function that runs when tenant application form is submitted
 * Automatically adds "Processed" column to the response sheet and sets new responses to "Pending Review"
 */
function onTenantApplicationSubmit(e) {
  try {
    console.log('Tenant application form submitted - looking for tenant application sheet...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let tenantSheet = null;
    
    // First, try to find "Tenant Application" sheet (if already renamed)
    tenantSheet = spreadsheet.getSheetByName('Tenant Application');
    
    // If not found, look for tenant application by checking headers
    if (!tenantSheet) {
      const sheets = spreadsheet.getSheets();
      for (let sheet of sheets) {
        const sheetName = sheet.getName();
        // Check if it's a Form Responses sheet
        if (sheetName.startsWith('Form Responses') && sheet.getLastColumn() > 0) {
          const headers = sheet.getRange(1, 1, 1, Math.min(10, sheet.getLastColumn())).getValues()[0];
          
          // Check if headers indicate this is a tenant application sheet
          const isTenantSheet = headers.some(header => 
            header && (
              header.toString().toLowerCase().includes('full name') ||
              header.toString().toLowerCase().includes('monthly income') ||
              header.toString().toLowerCase().includes('employment status') ||
              header.toString().toLowerCase().includes('current address')
            )
          );
          
          if (isTenantSheet) {
            tenantSheet = sheet;
            console.log(`Found tenant application sheet: ${sheetName}`);
            
            // Rename it now that we found it
            if (sheetName.startsWith('Form Responses')) {
              try {
                sheet.setName('Tenant Application');
                console.log(`✅ Renamed ${sheetName} to "Tenant Application"`);
              } catch (renameError) {
                console.log(`Could not rename sheet: ${renameError.message}`);
              }
            }
            break;
          }
        }
      }
    }
    
    if (!tenantSheet) {
      console.log('❌ Could not find tenant application sheet');
      // List available sheets for debugging
      const sheets = spreadsheet.getSheets();
      console.log('Available sheets:', sheets.map(s => s.getName()));
      return;
    }
    
    console.log('✅ Found tenant application sheet:', tenantSheet.getName());
    
    // Check if "Processed" column already exists
    const lastColumn = tenantSheet.getLastColumn();
    const headers = lastColumn > 0 ? tenantSheet.getRange(1, 1, 1, lastColumn).getValues()[0] : [];
    
    let processedColumnIndex = headers.findIndex(header => 
      header && header.toString().toLowerCase().includes('processed')
    );
    
    // If Processed column doesn't exist, create it
    if (processedColumnIndex === -1) {
      console.log('Creating Processed column with dropdown...');
      
      // Add "Processed" header
      processedColumnIndex = lastColumn;
      tenantSheet.getRange(1, processedColumnIndex + 1).setValue('Processed');
      
      // Format the header
      tenantSheet.getRange(1, processedColumnIndex + 1)
        .setBackground('#1c4587')
        .setFontColor('white')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      
      // Create dropdown validation rule
      const dropdownOptions = [
        'Pending Review',
        'Approved',
        'Rejected'
      ];
      
      const validationRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(dropdownOptions)
        .setAllowInvalid(false)
        .setHelpText('Select the processing status for this application')
        .build();
      
      // Apply dropdown to a large range (rows 2-500)
      const dropdownRange = tenantSheet.getRange(2, processedColumnIndex + 1, 500, 1);
      dropdownRange.setDataValidation(validationRule);
      
      // Set column width
      tenantSheet.setColumnWidth(processedColumnIndex + 1, 150);
      
      console.log('✅ Dropdown validation applied');
      
      // Set all existing rows to "Pending Review"
      const lastRow = tenantSheet.getLastRow();
      if (lastRow > 1) {
        const dataRange = tenantSheet.getRange(2, processedColumnIndex + 1, lastRow - 1, 1);
        const values = [];
        for (let i = 0; i < lastRow - 1; i++) {
          values.push(['Pending Review']);
        }
        dataRange.setValues(values);
        console.log(`Set ${lastRow - 1} existing rows to "Pending Review"`);
      }
    } else {
      // Column exists, just set the new response to "Pending Review"
      const lastRow = tenantSheet.getLastRow();
      if (lastRow > 1) {
        tenantSheet.getRange(lastRow, processedColumnIndex + 1).setValue('Pending Review');
        console.log(`✅ Set new response at row ${lastRow} to "Pending Review"`);
      }
    }
    
  } catch (error) {
    console.error('Error in onTenantApplicationSubmit trigger:', error);
  }
}

/**
 * Function to rename Form Response sheets to meaningful names
 * This runs on a delayed trigger after forms are created
 */
function renameFormResponseSheets() {
  try {
    console.log('Running delayed sheet renaming...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    for (let sheet of sheets) {
      const sheetName = sheet.getName();
      
      // Skip if not a Form Responses sheet
      if (!sheetName.startsWith('Form Responses')) {
        continue;
      }
      
      // Skip if no data (headers)
      if (sheet.getLastColumn() === 0) {
        continue;
      }
      
      console.log(`Analyzing sheet: ${sheetName}`);
      
      try {
        const headers = sheet.getRange(1, 1, 1, Math.min(10, sheet.getLastColumn())).getValues()[0];
        console.log(`Headers: ${headers.slice(0, 5).join(', ')}...`);
        
        let newName = null;
        
        // Check if it's a tenant application sheet
        const isTenantSheet = headers.some(header => 
          header && (
            header.toString().toLowerCase().includes('full name') ||
            header.toString().toLowerCase().includes('monthly income') ||
            header.toString().toLowerCase().includes('employment status')
          )
        );
        
        // Check if it's a move-out request sheet
        const isMoveOutSheet = headers.some(header => 
          header && (
            header.toString().toLowerCase().includes('tenant name') ||
            header.toString().toLowerCase().includes('planned move-out date') ||
            header.toString().toLowerCase().includes('forwarding address')
          )
        );
        
        // Check if it's a guest check-in sheet
        const isGuestSheet = headers.some(header => 
          header && (
            header.toString().toLowerCase().includes('guest name') ||
            header.toString().toLowerCase().includes('check-in date') ||
            header.toString().toLowerCase().includes('number of nights')
          )
        );
        
        if (isTenantSheet) {
          newName = 'Tenant Application';
        } else if (isMoveOutSheet) {
          newName = 'Move-Out Requests';
        } else if (isGuestSheet) {
          newName = 'Guest Check-Ins';
        }
        
        if (newName) {
          // Check if name already exists
          let finalName = newName;
          let counter = 1;
          while (spreadsheet.getSheetByName(finalName)) {
            finalName = `${newName} ${counter}`;
            counter++;
          }
          
          sheet.setName(finalName);
          console.log(`✅ Renamed "${sheetName}" to "${finalName}"`);
        }
        
      } catch (headerError) {
        console.log(`Could not analyze headers for ${sheetName}: ${headerError.message}`);
      }
    }
    
    console.log('✅ Sheet renaming completed');
    
  } catch (error) {
    console.error('Error in renameFormResponseSheets:', error);
  }
}

/**
 * Alternative function - manually add Processed column to Tenant Application sheet
 * Run this manually if the trigger approach doesn't work
 */
function manuallyAddProcessedColumn() {
  try {
    console.log('Manually adding Processed column...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const tenantSheet = spreadsheet.getSheetByName('Tenant Application');
    
    if (!tenantSheet) {
      console.log('❌ Could not find sheet named "Tenant Application"');
      // List all available sheets
      const sheets = spreadsheet.getSheets();
      console.log('Available sheets:', sheets.map(s => s.getName()));
      return 'Could not find "Tenant Application" sheet';
    }
    
    // Check if column already exists
    const lastColumn = tenantSheet.getLastColumn();
    const headers = lastColumn > 0 ? tenantSheet.getRange(1, 1, 1, lastColumn).getValues()[0] : [];
    
    const processedColumnExists = headers.some(header => 
      header && header.toString().toLowerCase().includes('processed')
    );
    
    if (processedColumnExists) {
      return 'Processed column already exists';
    }
    
    // Add header
    const processedColumnIndex = lastColumn + 1;
    tenantSheet.getRange(1, processedColumnIndex).setValue('Processed');
    
    // Format header
    tenantSheet.getRange(1, processedColumnIndex)
      .setBackground('#1c4587')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    // Create simplified dropdown - only 3 options
    const dropdownOptions = [
      'Pending Review',
      'Approved',
      'Rejected'
    ];
    
    const validationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(dropdownOptions)
      .setAllowInvalid(false)
      .setHelpText('Select the processing status for this application')
      .build();
    
    // Apply to large range
    const dropdownRange = tenantSheet.getRange(2, processedColumnIndex, 500, 1);
    dropdownRange.setDataValidation(validationRule);
    
    // Set column width
    tenantSheet.setColumnWidth(processedColumnIndex, 150);
    
    // Set existing rows to "Pending Review"
    const lastRow = tenantSheet.getLastRow();
    if (lastRow > 1) {
      const dataRange = tenantSheet.getRange(2, processedColumnIndex, lastRow - 1, 1);
      const values = [];
      for (let i = 0; i < lastRow - 1; i++) {
        values.push(['Pending Review']);
      }
      dataRange.setValues(values);
    }
    
    return `✅ Successfully added Processed column with simplified dropdown to "Tenant Application" sheet!`;
    
  } catch (error) {
    console.error('Error manually adding Processed column:', error);
    return 'Error: ' + error.message;
  }
}

/**
 * Daily check function that runs the appropriate email functions based on the date
 * This is triggered automatically by the system
 * 
 * NEW EMAIL SCHEDULE:
 * - Due date: 1st of each month
 * - Reminder: One week before (25th of previous month)
 * - Late notices: Day after due date (2nd) and every week until paid
 * - Payment status: Updated automatically every day
 */
function checkAndRunDailyTasks() {
  const today = new Date();
  const dayOfMonth = today.getDate();
  const currentHour = today.getHours();
  
  console.log(`Daily check running - Day: ${dayOfMonth}, Hour: ${currentHour}`);
  
  // Only run email functions at 9 AM to avoid multiple executions
  if (currentHour === 9) {
    
    // Update payment statuses every day (this runs first)
    console.log('Updating all tenant payment statuses...');
    EmailManager.updateAllPaymentStatuses();
    
    // **NEW: Sync calendar data daily at 9 AM**
    console.log('Daily calendar sync...');
    try {
      CalendarManager.syncAllTenantsToCalendar();
      CalendarManager.syncAllGuestsToCalendar();
      console.log('✅ Daily calendar sync completed');
    } catch (calendarError) {
      console.log('⚠️ Daily calendar sync failed:', calendarError.message);
    }
    
    // Send rent reminders one week before due date (25th of each month)
    if (dayOfMonth === 25) {
      console.log('Running rent reminders (one week before due date)...');
      EmailManager.sendRentReminders();
    }
    
    // Send monthly invoices on the 1st (due date)
    if (dayOfMonth === 1) {
      console.log('Running monthly invoices (due date)...');
      EmailManager.sendMonthlyInvoices();
    }
    
    // Send late payment alerts starting the day after due date (2nd) and every week
    // Now continues until payment status is Current
    if (dayOfMonth === 2 || dayOfMonth === 9 || dayOfMonth === 16 || dayOfMonth === 23) {
      console.log(`Running late payment alerts (day ${dayOfMonth})...`);
      EmailManager.sendLatePaymentAlerts();
    }
  }
}

/**
 * Wrapper functions for manual email sending (called from menu)
 */
function sendRentReminders() {
  return EmailManager.sendRentReminders();
}

function sendLatePaymentAlerts() {
  return EmailManager.sendLatePaymentAlerts();
}

function sendMonthlyInvoices() {
  return EmailManager.sendMonthlyInvoices();
}

/**
 * Update all tenant payment statuses (called from menu)
 */
function updateAllPaymentStatuses() {
  return EmailManager.updateAllPaymentStatuses();
}

/**
 * Send all payment alerts manually (Late, Overdue, and Partial payments)
 */
function sendAllPaymentAlerts() {
  return EmailManager.sendLatePaymentAlerts();
}

/**
 * Dashboard wrapper functions (called from menu)
 */
function createManagementDashboard() {
  return Dashboard.createManagementDashboard();
}

function createFinancialDashboard() {
  return Dashboard.createFinancialDashboard();
}

function refreshAllDashboards() {
  return Dashboard.refreshAllDashboards();
}

/**
 * Application Processing Panel wrapper function (called from menu)
 */
function showApplicationProcessingPanel() {
  return Panel.showApplicationProcessingPanel();
}

/**
 * Server-side function for processing applications (called from Panel HTML)
 */
function processApplication(applicationData) {
  return Panel.processApplication(applicationData);
}

/**
 * Server-side function for rejecting applications (called from Panel HTML)
 */
function markApplicationAsRejected(applicationData) {
  return Panel.markApplicationAsRejected(applicationData);
}

/**
 * Individual Booking Panel wrapper functions (called from menu)
 */
function showAvailabilityChecker() {
  return BookingPanels.showAvailabilityChecker();
}

function showNewBookingPanel() {
  return BookingPanels.showNewBookingPanel();
}

function showCheckInPanel() {
  return BookingPanels.showCheckInPanel();
}

function showOnlineReservationPanel() {
  return BookingPanels.showOnlineReservationPanel();
}

function processOnlineReservation(reservationData) {
  return BookingPanels.processOnlineReservation(reservationData);
}

function showCheckOutPanel() {
  return BookingPanels.showCheckOutPanel();
}

/**
 * Booking Manager wrapper functions (called from menu)
 */
function showBookingManagerPanel() {
  return BookingManager.showBookingManagerPanel();
}

function checkAvailability(availabilityData) {
  return BookingManager.checkAvailability(availabilityData);
}

function createNewBooking(bookingData) {
  return BookingManager.createNewBooking(bookingData);
}

function processCheckIn(checkInData) {
  return BookingManager.processCheckIn(checkInData);
}

function processCheckOut(checkOutData) {
  return BookingManager.processCheckOut(checkOutData);
}

function addTenantToCalendar(tenantData) {
  return BookingManager.addTenantToCalendar(tenantData);
}

/**
 * Calendar integration wrapper functions (called from menu)
 */
function syncAllTenantsToCalendar() {
  return CalendarManager.syncAllTenantsToCalendar();
}

function syncAllGuestsToCalendar() {
  return CalendarManager.syncAllGuestsToCalendar();
}

function syncAllToCalendar() {
  return CalendarManager.syncAllToCalendar();
}

function clearWhiteHouseCalendarEvents() {
  return CalendarManager.clearWhiteHouseCalendarEvents();
}

/**
 * Move-Out Processing Panel wrapper function (called from menu)
 */
function showMoveOutProcessingPanel() {
  return MoveOutPanel.showMoveOutProcessingPanel();
}

/**
 * Server-side function for processing move-outs (called from MoveOutPanel HTML)
 */
function processMoveOutRequest(moveOutData) {
  return MoveOutPanel.processMoveOutRequest(moveOutData);
}

/**
 * Payment Recording Panel wrapper function (called from menu)
 */
function showPaymentRecordingPanel() {
  return PaymentPanel.showPaymentRecordingPanel();
}

/**
 * Server-side function for recording payments (called from PaymentPanel HTML)
 */
function recordTenantPayment(paymentData) {
  return PaymentPanel.recordTenantPayment(paymentData);
}

/**
 * Budget Entry Panel wrapper function (called from menu)
 * NEW: Added for easy income/expense entry
 */
function showBudgetEntryPanel() {
  return BudgetPanel.showBudgetEntryPanel();
}

/**
 * Server-side function for adding budget entries (called from BudgetPanel HTML)
 * NEW: Added for budget entry functionality
 */
function addBudgetEntry(entryData) {
  return BudgetPanel.addBudgetEntry(entryData);
}

/**
 * Maintenance Request Panel wrapper function (called from menu)
 */
function showMaintenanceRequestPanel() {
  return MaintenancePanel.showMaintenanceRequestPanel();
}

/**
 * Server-side function for adding maintenance requests (called from MaintenancePanel HTML)
 */
function addMaintenanceRequest(requestData) {
  return MaintenancePanel.addMaintenanceRequest(requestData);
}

/**
 * Utility functions for testing and debugging
 */
function testRentReminder() {
  return EmailManager.testRentReminder();
}

function getTriggerInfo() {
  return TriggerManager.getTriggerInfo();
}
