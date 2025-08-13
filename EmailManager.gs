/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Email Manager - EmailManager.gs
 * 
 * This module handles all email communications with updated timing:
 * - Reminders: One week before due date (25th)
 * - Invoices: On due date (1st)
 * - Late alerts: Day after due date (2nd) and weekly (9th, 16th, 23rd)
 * 
 * UPDATES:
 * - Fixed rent amount to use negotiated price when available
 * - Added Zelle payment instructions to all emails
 * - Added partial payment email support
 * - Automatic payment status updates
 */

const EmailManager = {

  /**
   * Send rent reminders to all active tenants (one week before due date)
   */
  sendRentReminders() {
    try {
      console.log('Sending rent reminders...');
      
      const tenants = this._getActiveTenants();
      let sentCount = 0;
      
      tenants.forEach(tenant => {
        if (tenant.email && tenant.name) {
          const subject = `Rent Reminder - Room ${tenant.roomNumber}`;
          const body = this._createRentReminderEmail(tenant);
          
          try {
            GmailApp.sendEmail(tenant.email, subject, body);
            sentCount++;
            console.log(`Rent reminder sent to ${tenant.name} (${tenant.email})`);
          } catch (emailError) {
            console.error(`Failed to send reminder to ${tenant.email}:`, emailError);
          }
        }
      });
      
      console.log(`Rent reminders sent to ${sentCount} tenants`);
      return `Rent reminders sent to ${sentCount} tenants`;
    } catch (error) {
      console.error('Error sending rent reminders:', error);
      throw error;
    }
  },

  /**
   * Send late payment alerts to tenants with overdue payments
   * Now runs day after due date and weekly thereafter
   * UPDATED: Continues sending until payment status is Current, includes partial payments
   */
  sendLatePaymentAlerts() {
    try {
      console.log('Sending late payment alerts...');
      
      // First, update all tenant payment statuses automatically
      this._updateAllPaymentStatuses();
      
      const tenants = this._getActiveTenants();
      const currentDate = new Date();
      let sentCount = 0;
      
      tenants.forEach(tenant => {
        // Send alerts to anyone with Late, Overdue, or Partial status (not Current)
        if (this._shouldSendLateAlert(tenant) && tenant.email && tenant.name) {
          const subject = this._getEmailSubject(tenant);
          const body = this._createPaymentAlertEmail(tenant);
          
          try {
            GmailApp.sendEmail(tenant.email, subject, body);
            sentCount++;
            console.log(`Payment alert sent to ${tenant.name} (${tenant.email}) - Status: ${tenant.paymentStatus}`);
          } catch (emailError) {
            console.error(`Failed to send payment alert to ${tenant.email}:`, emailError);
          }
        }
      });
      
      console.log(`Payment alerts sent to ${sentCount} tenants`);
      return `Payment alerts sent to ${sentCount} tenants`;
    } catch (error) {
      console.error('Error sending payment alerts:', error);
      throw error;
    }
  },

  /**
   * Send monthly invoices to all active tenants (on due date)
   */
  sendMonthlyInvoices() {
    try {
      console.log('Sending monthly invoices...');
      
      const tenants = this._getActiveTenants();
      let sentCount = 0;
      
      tenants.forEach(tenant => {
        if (tenant.email && tenant.name) {
          const subject = `Monthly Rent Invoice - Room ${tenant.roomNumber}`;
          const body = this._createMonthlyInvoiceEmail(tenant);
          
          try {
            GmailApp.sendEmail(tenant.email, subject, body);
            sentCount++;
            console.log(`Monthly invoice sent to ${tenant.name} (${tenant.email})`);
          } catch (emailError) {
            console.error(`Failed to send invoice to ${tenant.email}:`, emailError);
          }
        }
      });
      
      console.log(`Monthly invoices sent to ${sentCount} tenants`);
      return `Monthly invoices sent to ${sentCount} tenants`;
    } catch (error) {
      console.error('Error sending monthly invoices:', error);
      throw error;
    }
  },

  /**
   * Test function to preview rent reminder email
   */
  testRentReminder() {
    const tenants = this._getActiveTenants();
    if (tenants.length > 0) {
      const testTenant = tenants[0];
      console.log('Testing rent reminder with tenant:', testTenant.name);
      
      const subject = `TEST - Rent Reminder - Room ${testTenant.roomNumber}`;
      const body = this._createRentReminderEmail(testTenant);
      
      console.log('Subject:', subject);
      console.log('Body:', body);
      
      return 'Test rent reminder created (check logs for content)';
    } else {
      return 'No active tenants found for testing';
    }
  },

  /**
   * Manual function to update all tenant payment statuses
   * Can be called from menu or manually
   */
  updateAllPaymentStatuses() {
    try {
      console.log('Manual payment status update triggered...');
      this._updateAllPaymentStatuses();
      return 'Payment statuses updated for all tenants based on current date and last payment dates.';
    } catch (error) {
      console.error('Error in manual payment status update:', error);
      throw error;
    }
  },

  /**
   * Get all active tenants from the Tenant sheet
   * UPDATED: Now properly gets negotiated rent when available
   * @private
   */
  _getActiveTenants() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TENANT);
    if (!sheet) {
      throw new Error('Tenant sheet not found');
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const tenants = [];
    
    // Find column indices
    const roomNumberCol = headers.indexOf('Room Number');
    const rentalPriceCol = headers.indexOf('Rental Price');
    const negotiatedPriceCol = headers.indexOf('Negotiated Price');
    const nameCol = headers.indexOf('Current Tenant Name');
    const emailCol = headers.indexOf('Tenant Email');
    const phoneCol = headers.indexOf('Tenant Phone');
    const statusCol = headers.indexOf('Room Status');
    const lastPaymentCol = headers.indexOf('Last Payment Date');
    const paymentStatusCol = headers.indexOf('Payment Status');
    
    // Process each row (skip header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Only include active tenants (not vacant rooms)
      if (row[statusCol] && row[statusCol].toString().toLowerCase() === 'occupied' && row[nameCol]) {
        
        // FIXED: Use negotiated price if available, otherwise fall back to rental price
        let rentAmount = row[rentalPriceCol];
        if (row[negotiatedPriceCol] && row[negotiatedPriceCol].toString().trim() !== '') {
          rentAmount = row[negotiatedPriceCol];
          console.log(`Using negotiated rent for ${row[nameCol]}: ${rentAmount} (original: ${row[rentalPriceCol]})`);
        } else {
          console.log(`Using rental price for ${row[nameCol]}: ${rentAmount}`);
        }
        
        tenants.push({
          roomNumber: row[roomNumberCol],
          rentalPrice: row[rentalPriceCol],
          negotiatedPrice: row[negotiatedPriceCol],
          rentAmount: rentAmount,  // This is the actual amount to use in emails
          name: row[nameCol],
          email: row[emailCol],
          phone: row[phoneCol],
          status: row[statusCol],
          lastPaymentDate: row[lastPaymentCol],
          paymentStatus: row[paymentStatusCol]
        });
      }
    }
    
    return tenants;
  },

  /**
   * Automatically update all tenant payment statuses based on current date
   * Called before sending late payment alerts
   * @private
   */
  _updateAllPaymentStatuses() {
    try {
      console.log('Updating all tenant payment statuses...');
      
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TENANT);
      if (!sheet) {
        console.error('Tenant sheet not found for status updates');
        return;
      }
      
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      
      // Find column indices
      const roomNumberCol = headers.indexOf('Room Number');
      const nameCol = headers.indexOf('Current Tenant Name');
      const statusCol = headers.indexOf('Room Status');
      const lastPaymentCol = headers.indexOf('Last Payment Date');
      const paymentStatusCol = headers.indexOf('Payment Status');
      
      if (paymentStatusCol === -1) {
        console.error('Payment Status column not found');
        return;
      }
      
      const today = new Date();
      const currentDay = today.getDate();
      const currentMonth = today.getMonth();
      const currentYear = today.getFullYear();
      
      // Calculate this month's due date (1st of current month)
      const thisMonthDueDate = new Date(currentYear, currentMonth, 1);
      
      let updatedCount = 0;
      
      // Process each tenant (skip header row)
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // Only process occupied rooms with tenants
        if (row[statusCol] && row[statusCol].toString().toLowerCase() === 'occupied' && row[nameCol]) {
          const tenantName = row[nameCol];
          const lastPaymentDate = row[lastPaymentCol];
          const currentPaymentStatus = row[paymentStatusCol] || '';
          
          // Skip if already marked as Current or Partial (don't override Partial status)
          if (currentPaymentStatus === 'Current' || currentPaymentStatus === 'Partial') {
            continue;
          }
          
          let newStatus = '';
          
          if (lastPaymentDate) {
            const lastPayment = new Date(lastPaymentDate);
            
            // Check if last payment was for this month (after due date)
            if (lastPayment >= thisMonthDueDate) {
              newStatus = 'Current';
            } else if (currentDay >= 8) { // 1 week after due date
              newStatus = 'Overdue';
            } else if (currentDay >= 2) { // 1 day after due date
              newStatus = 'Late';
            }
          } else {
            // No payment recorded
            if (currentDay >= 8) {
              newStatus = 'Overdue';
            } else if (currentDay >= 2) {
              newStatus = 'Late';
            }
          }
          
          // Update status if it changed
          if (newStatus && newStatus !== currentPaymentStatus) {
            sheet.getRange(i + 1, paymentStatusCol + 1).setValue(newStatus);
            console.log(`Updated ${tenantName} (Room ${row[roomNumberCol]}): ${currentPaymentStatus} → ${newStatus}`);
            updatedCount++;
          }
        }
      }
      
      console.log(`Updated payment status for ${updatedCount} tenants`);
      
    } catch (error) {
      console.error('Error updating payment statuses:', error);
    }
  },

  /**
   * Check if tenant should receive payment alert (late, overdue, or partial)
   * @private
   */
  _shouldSendLateAlert(tenant) {
    const paymentStatus = tenant.paymentStatus || '';
    return paymentStatus === 'Late' || paymentStatus === 'Overdue' || paymentStatus === 'Partial';
  },

  /**
   * Get appropriate email subject based on payment status
   * @private
   */
  _getEmailSubject(tenant) {
    const paymentStatus = tenant.paymentStatus || '';
    
    if (paymentStatus === 'Partial') {
      return `Partial Payment - Balance Due - Room ${tenant.roomNumber}`;
    } else {
      return `Payment Overdue Notice - Room ${tenant.roomNumber}`;
    }
  },

  /**
   * Create payment alert email (handles Late, Overdue, and Partial payments)
   * @private
   */
  _createPaymentAlertEmail(tenant) {
    const paymentStatus = tenant.paymentStatus || '';
    
    if (paymentStatus === 'Partial') {
      return this._createPartialPaymentEmail(tenant);
    } else {
      return this._createLatePaymentEmail(tenant);
    }
  },

  /**
   * Create rent reminder email content (sent one week before due date)
   * UPDATED: Uses correct rent amount and includes Zelle payment instructions
   * @private
   */
  _createRentReminderEmail(tenant) {
    const nextMonth = new Date();
    nextMonth.setMonth(nextMonth.getMonth() + 1);
    const monthName = nextMonth.toLocaleString('default', { month: 'long', year: 'numeric' });
    const rentAmount = tenant.rentAmount || 'N/A';  // Uses negotiated price when available
    
    return `
Dear ${tenant.name},

This is a friendly reminder that your rent payment for ${monthName} is due on the 1st of the month.

Room Details:
- Room Number: ${tenant.roomNumber}
- Monthly Rent: ${rentAmount}
- Due Date: ${monthName} 1st

You can send your rent via Zelle to our bookkeeper Akiko at bookkeeper@belvederefamily.com, or you may give the cash directly to the house manager.

Please ensure your payment is submitted by the due date. This is just a courtesy reminder to help you stay on track.

If you have already made your payment, please disregard this message.

Thank you for being a valued tenant!

Best regards,
${EMAIL_CONFIG.MANAGEMENT_TEAM}
${EMAIL_CONFIG.PROPERTY_NAME}
    `.trim();
  },

  /**
   * Create late payment email content (sent after due date)
   * UPDATED: Uses correct rent amount, Zelle instructions, and accurate timing
   * @private
   */
  _createLatePaymentEmail(tenant) {
    const currentMonth = new Date().toLocaleString('default', { month: 'long', year: 'numeric' });
    const rentAmount = tenant.rentAmount || 'N/A';  // Uses negotiated price when available
    const today = new Date();
    const dayOfMonth = today.getDate();
    
    // More accurate urgency messaging
    let urgencyLevel = '';
    let daysPastDue = dayOfMonth - 1; // Days past the 1st
    
    if (daysPastDue === 1) {
      urgencyLevel = 'Your rent payment is now 1 day overdue.';
    } else if (daysPastDue <= 7) {
      urgencyLevel = `Your rent payment is now ${daysPastDue} days overdue.`;
    } else if (daysPastDue <= 14) {
      urgencyLevel = `Your rent payment is now ${daysPastDue} days overdue (over 1 week late).`;
    } else if (daysPastDue <= 21) {
      urgencyLevel = `Your rent payment is now ${daysPastDue} days overdue (over 2 weeks late).`;
    } else {
      urgencyLevel = `Your rent payment is now ${daysPastDue} days overdue (over 3 weeks late).`;
    }
    
    return `
Dear ${tenant.name},

${urgencyLevel}

Room Details:
- Room Number: ${tenant.roomNumber}
- Monthly Rent: ${rentAmount}
- Amount Due: ${rentAmount}
- Original Due Date: ${currentMonth} 1st

You can send your rent via Zelle to our bookkeeper Akiko at bookkeeper@belvederefamily.com, or you may give the cash directly to the house manager.

Please submit your payment immediately to bring your account current.

If you have already made your payment or are experiencing financial difficulties, please contact us immediately to discuss the situation.

We value you as a tenant and want to work with you to resolve this matter promptly.

Best regards,
${EMAIL_CONFIG.MANAGEMENT_TEAM}
${EMAIL_CONFIG.PROPERTY_NAME}
    `.trim();
  },

  /**
   * Create monthly invoice email content (sent on due date)
   * UPDATED: Uses correct rent amount and includes Zelle payment instructions
   * @private
   */
  _createMonthlyInvoiceEmail(tenant) {
    const currentDate = new Date();
    const currentMonth = currentDate.toLocaleString('default', { month: 'long', year: 'numeric' });
    const rentAmount = tenant.rentAmount || 'N/A';  // Uses negotiated price when available
    const dueDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
    
    return `
Dear ${tenant.name},

Your monthly rent is due today. Please find your invoice details below:

RENT INVOICE - ${currentMonth}
Room Number: ${tenant.roomNumber}
Tenant: ${tenant.name}
Amount Due: ${rentAmount}
Due Date: ${dueDate.toLocaleDateString()}

You can send your rent via Zelle to our bookkeeper Akiko at bookkeeper@belvederefamily.com, or you may give the cash directly to the house manager.

Please ensure payment is submitted today to avoid your account becoming overdue.

If you have any questions about this invoice, please don't hesitate to contact us.

Thank you!

Best regards,
${EMAIL_CONFIG.MANAGEMENT_TEAM}
${EMAIL_CONFIG.PROPERTY_NAME}
    `.trim();
  },

  /**
   * Create partial payment email content (no specific amounts mentioned)
   * @private
   */
  _createPartialPaymentEmail(tenant) {
    const currentMonth = new Date().toLocaleString('default', { month: 'long', year: 'numeric' });
    const rentAmount = tenant.rentAmount || 'N/A';  // Uses negotiated price when available
    
    return `
Dear ${tenant.name},

We have received a partial payment for your ${currentMonth} rent. However, there is still a remaining balance on your account that needs to be paid.

Room Details:
- Room Number: ${tenant.roomNumber}
- Monthly Rent: ${rentAmount}
- Status: Partial payment received
- Balance Due: Remaining amount still owed

You can send your remaining balance via Zelle to our bookkeeper Akiko at bookkeeper@belvederefamily.com, or you may give the cash directly to the house manager.

Please submit the remaining balance as soon as possible to bring your account current and avoid additional late fees.

If you have any questions about your payment or remaining balance, please contact us immediately.

Thank you for your partial payment, and we look forward to receiving the remaining amount soon.

Best regards,
${EMAIL_CONFIG.MANAGEMENT_TEAM}
${EMAIL_CONFIG.PROPERTY_NAME}
    `.trim();
  }
};
