/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Calendar Manager - CalendarManager.gs
 * 
 * This module handles Google Calendar integration for both tenants and guests
 */

const CalendarManager = {

  /**
   * Sync all tenants to Google Calendar
   */
  syncAllTenantsToCalendar() {
    try {
      console.log('Syncing all tenants to Google Calendar...');
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const tenantSheet = spreadsheet.getSheetByName(SHEET_NAMES.TENANT);
      
      if (!tenantSheet || tenantSheet.getLastRow() <= 1) {
        return 'No tenants found to sync';
      }
      
      const data = tenantSheet.getDataRange().getValues();
      const headers = data[0];
      let syncedCount = 0;
      
      const roomNumberCol = headers.indexOf('Room Number');
      const nameCol = headers.indexOf('Current Tenant Name');
      const emailCol = headers.indexOf('Tenant Email');
      const phoneCol = headers.indexOf('Tenant Phone');
      const moveInCol = headers.indexOf('Move-In Date');
      const leaseEndCol = headers.indexOf('Lease End Date');
      const statusCol = headers.indexOf('Room Status');
      const rentCol = headers.indexOf('Negotiated Price') !== -1 ? 
        headers.indexOf('Negotiated Price') : headers.indexOf('Rental Price');
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // Only sync occupied rooms with tenants
        if (row[statusCol] === 'Occupied' && row[nameCol] && row[moveInCol]) {
          try {
            this._addTenantToCalendar({
              tenantName: row[nameCol],
              email: row[emailCol] || '',
              phone: row[phoneCol] || '',
              roomNumber: row[roomNumberCol],
              moveInDate: row[moveInCol],
              leaseEndDate: row[leaseEndCol] || '',
              rentAmount: row[rentCol] || ''
            });
            syncedCount++;
          } catch (error) {
            console.error(`Error syncing tenant ${row[nameCol]}:`, error);
          }
        }
      }
      
      console.log(`Synced ${syncedCount} tenants to Google Calendar`);
      return `✅ Synced ${syncedCount} tenants to Google Calendar`;
      
    } catch (error) {
      console.error('Error syncing tenants to calendar:', error);
      throw new Error('Failed to sync tenants to calendar: ' + error.message);
    }
  },

  /**
   * Sync all guest bookings to Google Calendar
   */
  syncAllGuestsToCalendar() {
    try {
      console.log('Syncing all guest bookings to Google Calendar...');
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const guestSheet = spreadsheet.getSheetByName(SHEET_NAMES.GUEST_ROOMS);
      
      if (!guestSheet || guestSheet.getLastRow() <= 1) {
        return 'No guest bookings found to sync';
      }
      
      const data = guestSheet.getDataRange().getValues();
      const headers = data[0];
      let syncedCount = 0;
      
      const roomNumberCol = headers.indexOf('Room Number');
      const currentGuestCol = headers.indexOf('Current Guest');
      const checkInCol = headers.indexOf('Check-In Date');
      const checkOutCol = headers.indexOf('Check-Out Date');
      const statusCol = headers.indexOf('Status');
      const bookingStatusCol = headers.indexOf('Booking Status');
      const numberOfGuestsCol = headers.indexOf('Number of Guests');
      const purposeCol = headers.indexOf('Purpose of Visit');
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // Only sync active bookings
        if (row[currentGuestCol] && row[checkInCol] && row[checkOutCol] &&
            (row[statusCol] === 'Reserved' || row[statusCol] === 'Occupied') &&
            row[bookingStatusCol] !== 'Checked-Out' && row[bookingStatusCol] !== 'Cancelled') {
          
          try {
            this._addGuestToCalendar({
              guestName: row[currentGuestCol],
              roomNumber: row[roomNumberCol],
              checkInDate: row[checkInCol],
              checkOutDate: row[checkOutCol],
              numberOfGuests: row[numberOfGuestsCol] || '1',
              purposeOfVisit: row[purposeCol] || '',
              status: row[statusCol]
            });
            syncedCount++;
          } catch (error) {
            console.error(`Error syncing guest ${row[currentGuestCol]}:`, error);
          }
        }
      }
      
      console.log(`Synced ${syncedCount} guest bookings to Google Calendar`);
      return `✅ Synced ${syncedCount} guest bookings to Google Calendar`;
      
    } catch (error) {
      console.error('Error syncing guests to calendar:', error);
      throw new Error('Failed to sync guests to calendar: ' + error.message);
    }
  },

  /**
   * Add individual tenant to Google Calendar
   * @private
   */
  _addTenantToCalendar(tenantData) {
    try {
      const calendar = CalendarApp.getDefaultCalendar();
      
      const startDate = new Date(tenantData.moveInDate);
      let endDate;
      
      if (tenantData.leaseEndDate) {
        endDate = new Date(tenantData.leaseEndDate);
      } else {
        // If no lease end date, set to 1 year from move-in
        endDate = new Date(startDate);
        endDate.setFullYear(endDate.getFullYear() + 1);
      }
      
      const title = `🏠 Tenant: ${tenantData.tenantName} - Room ${tenantData.roomNumber}`;
      const description = `White House Tenant Lease\n\n` +
                         `Tenant: ${tenantData.tenantName}\n` +
                         `Room: ${tenantData.roomNumber}\n` +
                         `Email: ${tenantData.email}\n` +
                         `Phone: ${tenantData.phone}\n` +
                         `Rent: ${tenantData.rentAmount}\n` +
                         `Move-in: ${startDate.toLocaleDateString()}\n` +
                         `Lease End: ${endDate.toLocaleDateString()}`;
      
      // Create all-day event
      calendar.createAllDayEvent(title, startDate, endDate, {
        description: description,
        location: `White House - Room ${tenantData.roomNumber}`
      });
      
      console.log(`Added tenant to calendar: ${tenantData.tenantName}`);
      
    } catch (error) {
      console.error('Error adding tenant to calendar:', error);
      throw error;
    }
  },

  /**
   * Add individual guest to Google Calendar
   * @private
   */
  _addGuestToCalendar(guestData) {
    try {
      const calendar = CalendarApp.getDefaultCalendar();
      
      const startDate = new Date(guestData.checkInDate);
      const endDate = new Date(guestData.checkOutDate);
      
      const title = `🏨 Guest: ${guestData.guestName} - Room ${guestData.roomNumber}`;
      const description = `White House Guest Booking\n\n` +
                         `Guest: ${guestData.guestName}\n` +
                         `Room: ${guestData.roomNumber}\n` +
                         `Check-in: ${startDate.toLocaleDateString()}\n` +
                         `Check-out: ${endDate.toLocaleDateString()}\n` +
                         `Guests: ${guestData.numberOfGuests}\n` +
                         `Purpose: ${guestData.purposeOfVisit}\n` +
                         `Status: ${guestData.status}`;
      
      // Create all-day event
      calendar.createAllDayEvent(title, startDate, endDate, {
        description: description,
        location: `White House - Room ${guestData.roomNumber}`
      });
      
      console.log(`Added guest to calendar: ${guestData.guestName}`);
      
    } catch (error) {
      console.error('Error adding guest to calendar:', error);
      throw error;
    }
  },

  /**
   * Clear all White House events from calendar
   */
  clearWhiteHouseCalendarEvents() {
    try {
      console.log('Clearing White House events from calendar...');
      
      const calendar = CalendarApp.getDefaultCalendar();
      const now = new Date();
      const oneYearAgo = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate());
      const oneYearFromNow = new Date(now.getFullYear() + 1, now.getMonth(), now.getDate());
      
      const events = calendar.getEvents(oneYearAgo, oneYearFromNow);
      let deletedCount = 0;
      
      events.forEach(event => {
        const title = event.getTitle();
        if (title.includes('🏠 Tenant:') || title.includes('🏨 Guest:') || 
            title.includes('White House')) {
          event.deleteEvent();
          deletedCount++;
        }
      });
      
      console.log(`Deleted ${deletedCount} White House events from calendar`);
      return `✅ Cleared ${deletedCount} White House events from calendar`;
      
    } catch (error) {
      console.error('Error clearing calendar events:', error);
      throw new Error('Failed to clear calendar events: ' + error.message);
    }
  }
};

/**
 * Wrapper functions for menu integration
 */
function syncAllTenantsToCalendar() {
  return CalendarManager.syncAllTenantsToCalendar();
}

function syncAllGuestsToCalendar() {
  return CalendarManager.syncAllGuestsToCalendar();
}

function syncAllToCalendar() {
  try {
    const tenantResult = CalendarManager.syncAllTenantsToCalendar();
    const guestResult = CalendarManager.syncAllGuestsToCalendar();
    return `${tenantResult}\n${guestResult}`;
  } catch (error) {
    console.error('Error syncing all to calendar:', error);
    return 'Error syncing to calendar: ' + error.message;
  }
}

function clearWhiteHouseCalendarEvents() {
  return CalendarManager.clearWhiteHouseCalendarEvents();
}
