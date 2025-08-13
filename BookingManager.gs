/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Booking Manager - BookingManager.gs
 * 
 * This module provides comprehensive booking management for guest rooms
 * with availability checking, booking creation, check-in/out processing,
 * and Google Calendar integration.
 */

const BookingManager = {

  /**
   * Show the main booking manager panel
   */
  showBookingManagerPanel() {
    try {
      console.log('Opening Booking Manager Panel...');
      
      const html = this._generateBookingManagerHTML();
      const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(1000)
        .setHeight(800)
        .setTitle('🏨 Guest Booking Manager');
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Guest Booking Manager');
      
    } catch (error) {
      console.error('Error showing booking manager:', error);
      SpreadsheetApp.getUi().alert('Error', 'Failed to load booking manager: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  },

  /**
   * Check room availability for given date range
   */
  checkAvailability(availabilityData) {
    try {
      console.log('Checking availability:', availabilityData);
      
      const data = JSON.parse(availabilityData);
      const startDate = new Date(data.startDate);
      const endDate = new Date(data.endDate);
      
      // Get all guest rooms
      const availableRooms = this._getAvailableRooms(startDate, endDate);
      
      return JSON.stringify({
        success: true,
        rooms: availableRooms,
        nights: this._calculateNights(startDate, endDate),
        dateRange: `${startDate.toLocaleDateString()} - ${endDate.toLocaleDateString()}`
      });
      
    } catch (error) {
      console.error('Error checking availability:', error);
      return JSON.stringify({
        success: false,
        error: error.message
      });
    }
  },

  /**
   * Create new booking
   */
  createNewBooking(bookingData) {
    try {
      console.log('Creating new booking:', bookingData);
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const guestSheet = spreadsheet.getSheetByName(SHEET_NAMES.GUEST_ROOMS);
      
      if (!guestSheet) {
        throw new Error('Guest Rooms sheet not found');
      }
      
      const data = JSON.parse(bookingData);
      
      // Check for conflicts one more time
      const conflicts = this._checkDateConflicts(data.roomNumber, data.checkInDate, data.checkOutDate);
      if (conflicts.length > 0) {
        throw new Error(`Room ${data.roomNumber} is not available for selected dates`);
      }
      
      // Generate booking ID and calculate totals
      const bookingId = this._generateBookingId();
      const nights = this._calculateNights(new Date(data.checkInDate), new Date(data.checkOutDate));
      const totalAmount = parseFloat(data.dailyRate) * nights;
      
      // Create booking row with simplified structure
      const bookingRow = [
        bookingId,                           // Booking ID
        data.roomNumber,                     // Room Number
        data.roomName || 'Guest Room',      // Room Name
        data.roomType || 'Standard',        // Room Type
        `${parseFloat(data.dailyRate).toFixed(2)}`, // Daily Rate
        data.checkInDate,                    // Check-In Date
        data.checkOutDate,                   // Check-Out Date
        nights.toString(),                   // Number of Nights
        data.numberOfGuests,                 // Number of Guests
        data.guestName,                      // Current Guest
        data.guestEmail || '',               // Guest Email
        data.guestPhone || '',               // Guest Phone
        data.purposeOfVisit || '',           // Purpose of Visit
        `${totalAmount.toFixed(2)}`,        // Total Amount
        data.paymentStatus || 'Pending',     // Payment Status
        'Reserved',                          // Booking Status
        data.bookingSource || 'Direct',      // Source
        `Booking created on ${new Date().toLocaleDateString()}` // Notes
      ];
      
      // Add to guest rooms sheet
      const lastRow = guestSheet.getLastRow();
      guestSheet.getRange(lastRow + 1, 1, 1, bookingRow.length).setValues([bookingRow]);
      
      // Add to Google Calendar
      this._addToGoogleCalendar({
        type: 'guest',
        title: `Guest: ${data.guestName}`,
        startDate: data.checkInDate,
        endDate: data.checkOutDate,
        room: data.roomNumber,
        details: `Guest stay - Room ${data.roomNumber}\nGuest: ${data.guestName}\nPhone: ${data.guestPhone}\nGuests: ${data.numberOfGuests}`
      });
      
      console.log(`Booking ${bookingId} created successfully`);
      return `✅ Booking ${bookingId} created successfully for ${data.guestName} in Room ${data.roomNumber}`;
      
    } catch (error) {
      console.error('Error creating booking:', error);
      throw new Error('Failed to create booking: ' + error.message);
    }
  },

  /**
   * Process check-in
   */
  processCheckIn(checkInData) {
    try {
      console.log('Processing check-in:', checkInData);
      
      const data = JSON.parse(checkInData);
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const guestSheet = spreadsheet.getSheetByName(SHEET_NAMES.GUEST_ROOMS);
      
      // Find and update the booking
      const rowIndex = this._findBookingRow(guestSheet, data.bookingId);
      if (rowIndex === -1) {
        throw new Error(`Booking ${data.bookingId} not found`);
      }
      
      // Update booking status with simplified structure
      const headers = guestSheet.getRange(1, 1, 1, guestSheet.getLastColumn()).getValues()[0];
      const bookingStatusCol = headers.indexOf('Booking Status');
      const notesCol = headers.indexOf('Notes');
      
      if (bookingStatusCol !== -1) {
        guestSheet.getRange(rowIndex, bookingStatusCol + 1).setValue('Checked-In');
      }
      if (notesCol !== -1) {
        const currentNotes = guestSheet.getRange(rowIndex, notesCol + 1).getValue();
        const newNotes = `${currentNotes}\nChecked in on ${new Date().toLocaleDateString()}`;
        guestSheet.getRange(rowIndex, notesCol + 1).setValue(newNotes);
      }
      
      return `✅ Check-in completed for booking ${data.bookingId}`;
      
    } catch (error) {
      console.error('Error processing check-in:', error);
      throw new Error('Failed to process check-in: ' + error.message);
    }
  },

  /**
   * Process check-out
   */
  processCheckOut(checkOutData) {
    try {
      console.log('Processing check-out:', checkOutData);
      
      const data = JSON.parse(checkOutData);
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const guestSheet = spreadsheet.getSheetByName(SHEET_NAMES.GUEST_ROOMS);
      
      // Find and update the booking
      const rowIndex = this._findBookingRow(guestSheet, data.bookingId);
      if (rowIndex === -1) {
        throw new Error(`Booking ${data.bookingId} not found`);
      }
      
      // Update booking status for checkout with simplified structure
      const headers = guestSheet.getRange(1, 1, 1, guestSheet.getLastColumn()).getValues()[0];
      const bookingStatusCol = headers.indexOf('Booking Status');
      const notesCol = headers.indexOf('Notes');
      const currentGuestCol = headers.indexOf('Current Guest');
      
      if (bookingStatusCol !== -1) {
        guestSheet.getRange(rowIndex, bookingStatusCol + 1).setValue('Checked-Out');
      }
      if (currentGuestCol !== -1) {
        guestSheet.getRange(rowIndex, currentGuestCol + 1).setValue('');
      }
      if (notesCol !== -1) {
        const currentNotes = guestSheet.getRange(rowIndex, notesCol + 1).getValue();
        const newNotes = `${currentNotes}\nChecked out on ${new Date().toLocaleDateString()}`;
        guestSheet.getRange(rowIndex, notesCol + 1).setValue(newNotes);
      }
      
      return `✅ Check-out completed for booking ${data.bookingId}`;
      
    } catch (error) {
      console.error('Error processing check-out:', error);
      throw new Error('Failed to process check-out: ' + error.message);
    }
  },

  /**
   * Add tenant move-in to Google Calendar
   */
  addTenantToCalendar(tenantData) {
    try {
      const data = JSON.parse(tenantData);
      
      this._addToGoogleCalendar({
        type: 'tenant',
        title: `Tenant: ${data.tenantName}`,
        startDate: data.moveInDate,
        endDate: data.leaseEndDate,
        room: data.roomNumber,
        details: `Tenant lease - Room ${data.roomNumber}\nTenant: ${data.tenantName}\nEmail: ${data.email}\nPhone: ${data.phone}\nRent: ${data.rentAmount}`
      });
      
      return `✅ Tenant ${data.tenantName} added to Google Calendar`;
      
    } catch (error) {
      console.error('Error adding tenant to calendar:', error);
      throw new Error('Failed to add tenant to calendar: ' + error.message);
    }
  },

  /**
   * Get available rooms for date range
   * @private
   */
  _getAvailableRooms(startDate, endDate) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const guestSheet = spreadsheet.getSheetByName(SHEET_NAMES.GUEST_ROOMS);
      
      if (!guestSheet || guestSheet.getLastRow() <= 1) {
        return [];
      }
      
      const data = guestSheet.getDataRange().getValues();
      const headers = data[0];
      const availableRooms = [];
      
      // Get unique rooms from the sheet
      const rooms = new Map();
      
      const roomNumberCol = headers.indexOf('Room Number');
      const roomNameCol = headers.indexOf('Room Name');
      const roomTypeCol = headers.indexOf('Room Type');
      const dailyRateCol = headers.indexOf('Daily Rate');
      const bookingStatusCol = headers.indexOf('Booking Status');
      const checkInCol = headers.indexOf('Check-In Date');
      const checkOutCol = headers.indexOf('Check-Out Date');
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const roomNumber = row[roomNumberCol];
        
        if (roomNumber) {
          // Check if room exists in our map
          if (!rooms.has(roomNumber)) {
            rooms.set(roomNumber, {
              roomNumber: roomNumber,
              roomName: row[roomNameCol] || 'Guest Room',
              roomType: row[roomTypeCol] || 'Standard',
              dailyRate: row[dailyRateCol] || '$75',
              conflicts: []
            });
          }
          
          // Check for conflicts
          const roomData = rooms.get(roomNumber);
          const conflicts = this._checkDateConflicts(roomNumber, startDate, endDate);
          roomData.conflicts = conflicts;
        }
      }
      
      // Convert map to array and filter available rooms
      rooms.forEach((room, roomNumber) => {
        const isAvailable = room.conflicts.length === 0;
        
        availableRooms.push({
          roomNumber: room.roomNumber,
          roomName: room.roomName,
          roomType: room.roomType,
          dailyRate: room.dailyRate,
          available: isAvailable,
          status: isAvailable ? 'Available' : 'Occupied',
          conflicts: room.conflicts
        });
      });
      
      return availableRooms;
      
    } catch (error) {
      console.error('Error getting available rooms:', error);
      return [];
    }
  },

  /**
   * Check for date conflicts
   * @private
   */
  _checkDateConflicts(roomNumber, startDate, endDate) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const guestSheet = spreadsheet.getSheetByName(SHEET_NAMES.GUEST_ROOMS);
      
      if (!guestSheet || guestSheet.getLastRow() <= 1) {
        return [];
      }
      
      const data = guestSheet.getDataRange().getValues();
      const headers = data[0];
      const conflicts = [];
      
      const roomNumberCol = headers.indexOf('Room Number');
      const checkInCol = headers.indexOf('Check-In Date');
      const checkOutCol = headers.indexOf('Check-Out Date');
      const statusCol = headers.indexOf('Status');
      const currentGuestCol = headers.indexOf('Current Guest');
      
      const newStart = new Date(startDate);
      const newEnd = new Date(endDate);
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        if (row[roomNumberCol] === roomNumber && 
            row[checkInCol] && row[checkOutCol] &&
            row[bookingStatusCol] !== 'Checked-Out' && row[bookingStatusCol] !== 'Cancelled') {
          
          const existingStart = new Date(row[checkInCol]);
          const existingEnd = new Date(row[checkOutCol]);
          
          // Check for overlap
          if (newStart < existingEnd && newEnd > existingStart) {
            conflicts.push({
              guest: row[currentGuestCol] || 'Unknown',
              checkIn: existingStart.toLocaleDateString(),
              checkOut: existingEnd.toLocaleDateString(),
              status: row[statusCol]
            });
          }
        }
      }
      
      return conflicts;
      
    } catch (error) {
      console.error('Error checking conflicts:', error);
      return [];
    }
  },

  /**
   * Add event to Google Calendar
   * @private
   */
  _addToGoogleCalendar(eventData) {
    try {
      const calendar = CalendarApp.getDefaultCalendar();
      
      const startDate = new Date(eventData.startDate);
      const endDate = new Date(eventData.endDate);
      
      let title, description;
      
      if (eventData.type === 'guest') {
        title = `🏨 ${eventData.title} - Room ${eventData.room}`;
        description = `${eventData.details}\n\nWhite House Guest Booking`;
      } else {
        title = `🏠 ${eventData.title} - Room ${eventData.room}`;
        description = `${eventData.details}\n\nWhite House Tenant Lease`;
      }
      
      // For multi-day events, create all-day events
      if (eventData.type === 'guest') {
        calendar.createAllDayEvent(title, startDate, endDate, {
          description: description,
          location: `White House - Room ${eventData.room}`
        });
      } else {
        // For tenants, create a longer-term event
        calendar.createAllDayEvent(title, startDate, endDate, {
          description: description,
          location: `White House - Room ${eventData.room}`
        });
      }
      
      console.log(`Added to Google Calendar: ${title}`);
      
    } catch (error) {
      console.error('Error adding to Google Calendar:', error);
      // Don't throw error here - calendar is optional
    }
  },

  /**
   * Generate unique booking ID
   * @private
   */
  _generateBookingId() {
    const prefix = 'BK';
    const timestamp = Date.now().toString().slice(-6);
    const random = Math.floor(Math.random() * 100).toString().padStart(2, '0');
    return `${prefix}${timestamp}${random}`;
  },

  /**
   * Calculate number of nights
   * @private
   */
  _calculateNights(startDate, endDate) {
    const timeDiff = endDate.getTime() - startDate.getTime();
    return Math.ceil(timeDiff / (1000 * 3600 * 24));
  },

  /**
   * Find booking row by booking ID
   * @private
   */
  _findBookingRow(sheet, bookingId) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const bookingIdCol = headers.indexOf('Booking ID');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][bookingIdCol] === bookingId) {
        return i + 1; // Return 1-based row index
      }
    }
    
    return -1;
  },

  /**
   * Generate HTML for booking manager panel
   * @private
   */
  _generateBookingManagerHTML() {
    return `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background: #1c4587; color: white; padding: 15px; margin: -20px -20px 20px -20px; }
        .tab-container { margin-bottom: 20px; }
        .tab-buttons { display: flex; border-bottom: 2px solid #ddd; }
        .tab-button { 
            background: #f8f9fa; 
            border: none; 
            padding: 12px 20px; 
            cursor: pointer; 
            border-bottom: 2px solid transparent;
            margin-right: 5px;
        }
        .tab-button.active { 
            background: #1c4587; 
            color: white; 
            border-bottom-color: #1c4587;
        }
        .tab-content { 
            display: none; 
            background: #fff; 
            padding: 20px; 
            border: 1px solid #ddd; 
            border-top: none;
        }
        .tab-content.active { display: block; }
        .form-group { margin: 20px 0; }
        .form-group label { display: block; font-weight: bold; margin-bottom: 10px; }
        .form-group input, .form-group select, .form-group textarea { 
            width: calc(100% - 24px); 
            padding: 12px; 
            border: 1px solid #ccc; 
            border-radius: 4px; 
            font-size: 14px;
            margin-bottom: 8px;
            box-sizing: border-box;
        }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }
        .form-row-three { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 15px; margin-bottom: 20px; }
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
        .btn-success { background: #22803c; }
        .btn-success:hover { background: #1a6b30; }
        .btn-danger { background: #dc3545; }
        .btn-danger:hover { background: #c82333; }
        .status { padding: 10px; margin: 10px 0; border-radius: 4px; }
        .status.success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .status.error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .availability-results { margin-top: 20px; }
        .room-card { 
            border: 1px solid #ddd; 
            border-radius: 8px; 
            padding: 15px; 
            margin: 10px 0; 
            background: #f9f9f9;
        }
        .room-available { border-left: 4px solid #22803c; }
        .room-occupied { border-left: 4px solid #dc3545; }
        .booking-list { max-height: 400px; overflow-y: auto; }
        .booking-item { 
            border: 1px solid #ddd; 
            border-radius: 8px; 
            padding: 15px; 
            margin: 10px 0; 
            background: #fff;
        }
    </style>
</head>
<body>
    <div class="header">
        <h2>🏨 Guest Booking Manager</h2>
        <p>Manage guest bookings, check availability, and process check-ins/outs</p>
    </div>
    
    <div id="status" class="status" style="display: none;"></div>
    
    <div class="tab-container">
        <div class="tab-buttons">
            <button class="tab-button active" onclick="showTab('availability')">⚡ Check Availability</button>
            <button class="tab-button" onclick="showTab('booking')">📝 Create Booking</button>
            <button class="tab-button" onclick="showTab('checkin')">✅ Check-In</button>
            <button class="tab-button" onclick="showTab('checkout')">🚪 Check-Out</button>
        </div>
        
        <!-- Availability Tab -->
        <div id="availability-tab" class="tab-content active">
            <h3>⚡ Check Room Availability</h3>
            <div class="form-row">
                <div class="form-group">
                    <label>Check-In Date:</label>
                    <input type="date" id="avail-checkin" required>
                </div>
                <div class="form-group">
                    <label>Check-Out Date:</label>
                    <input type="date" id="avail-checkout" required>
                </div>
            </div>
            <button class="btn" onclick="checkAvailability()">🔍 Check Availability</button>
            
            <div id="availability-results" class="availability-results"></div>
        </div>
        
        <!-- Create Booking Tab -->
        <div id="booking-tab" class="tab-content">
            <h3>📝 Create New Booking</h3>
            <div class="form-row">
                <div class="form-group">
                    <label>Guest Name *:</label>
                    <input type="text" id="guest-name" required>
                </div>
                <div class="form-group">
                    <label>Guest Email:</label>
                    <input type="email" id="guest-email">
                </div>
            </div>
            <div class="form-row">
                <div class="form-group">
                    <label>Guest Phone:</label>
                    <input type="tel" id="guest-phone">
                </div>
                <div class="form-group">
                    <label>Number of Guests:</label>
                    <input type="number" id="guest-count" min="1" value="1">
                </div>
            </div>
            <div class="form-row">
                <div class="form-group">
                    <label>Check-In Date *:</label>
                    <input type="date" id="booking-checkin" required>
                </div>
                <div class="form-group">
                    <label>Check-Out Date *:</label>
                    <input type="date" id="booking-checkout" required>
                </div>
            </div>
            <div class="form-row">
                <div class="form-group">
                    <label>Room Number *:</label>
                    <select id="booking-room" required>
                        <option value="">Select room...</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>Daily Rate:</label>
                    <input type="number" id="daily-rate" step="0.01" min="0">
                </div>
            </div>
            <div class="form-group">
                <label>Purpose of Visit:</label>
                <input type="text" id="purpose-visit" placeholder="Business, vacation, etc.">
            </div>
            <div class="form-group">
                <label>Special Requests:</label>
                <textarea id="special-requests" placeholder="Any special requests or notes..."></textarea>
            </div>
            <button class="btn btn-success" onclick="createBooking()">📝 Create Booking</button>
        </div>
        
        <!-- Check-In Tab -->
        <div id="checkin-tab" class="tab-content">
            <h3>✅ Process Check-In</h3>
            <div class="form-group">
                <label>Booking ID:</label>
                <input type="text" id="checkin-booking-id" placeholder="e.g., BK123456">
            </div>
            <button class="btn btn-success" onclick="processCheckIn()">✅ Process Check-In</button>
        </div>
        
        <!-- Check-Out Tab -->
        <div id="checkout-tab" class="tab-content">
            <h3>🚪 Process Check-Out</h3>
            <div class="form-group">
                <label>Booking ID:</label>
                <input type="text" id="checkout-booking-id" placeholder="e.g., BK123456">
            </div>
            <button class="btn btn-danger" onclick="processCheckOut()">🚪 Process Check-Out</button>
        </div>
    </div>
    
    <script>
        // Set default dates
        document.addEventListener('DOMContentLoaded', function() {
            const today = new Date();
            const tomorrow = new Date(today);
            tomorrow.setDate(tomorrow.getDate() + 1);
            
            document.getElementById('avail-checkin').value = today.toISOString().split('T')[0];
            document.getElementById('avail-checkout').value = tomorrow.toISOString().split('T')[0];
            document.getElementById('booking-checkin').value = today.toISOString().split('T')[0];
            document.getElementById('booking-checkout').value = tomorrow.toISOString().split('T')[0];
        });
        
        function showTab(tabName) {
            // Hide all tabs
            const tabs = document.querySelectorAll('.tab-content');
            const buttons = document.querySelectorAll('.tab-button');
            
            tabs.forEach(tab => tab.classList.remove('active'));
            buttons.forEach(btn => btn.classList.remove('active'));
            
            // Show selected tab
            document.getElementById(tabName + '-tab').classList.add('active');
            
            // Find and activate the correct button
            const targetButton = document.querySelector('[onclick="showTab(\'' + tabName + '\')"]');
            if (targetButton) {
                targetButton.classList.add('active');
            }
        }
        
        function checkAvailability() {
            const checkIn = document.getElementById('avail-checkin').value;
            const checkOut = document.getElementById('avail-checkout').value;
            
            if (!checkIn || !checkOut) {
                showStatus('Please select both check-in and check-out dates.', 'error');
                return;
            }
            
            if (new Date(checkIn) >= new Date(checkOut)) {
                showStatus('Check-out date must be after check-in date.', 'error');
                return;
            }
            
            showStatus('Checking availability...', 'success');
            
            google.script.run
                .withSuccessHandler(function(result) {
                    const data = JSON.parse(result);
                    if (data.success) {
                        displayAvailability(data);
                    } else {
                        showStatus('Error: ' + data.error, 'error');
                    }
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, 'error');
                })
                .checkAvailability(JSON.stringify({
                    startDate: checkIn,
                    endDate: checkOut
                }));
        }
        
        function displayAvailability(data) {
            const resultsDiv = document.getElementById('availability-results');
            let html = '<h4>Availability Results for ' + data.dateRange + ' (' + data.nights + ' nights)</h4>';
            
            if (data.rooms.length === 0) {
                html += '<p>No rooms found in the system.</p>';
            } else {
                data.rooms.forEach(room => {
                    const cardClass = room.available ? 'room-available' : 'room-occupied';
                    const statusIcon = room.available ? '✅' : '❌';
                    const statusText = room.available ? 'Available' : 'Occupied';
                    
                    html += '<div class="room-card ' + cardClass + '">';
                    html += '<h5>' + statusIcon + ' Room ' + room.roomNumber + ' - ' + room.roomName + '</h5>';
                    html += '<p><strong>Type:</strong> ' + room.roomType + '</p>';
                    html += '<p><strong>Rate:</strong> ' + room.dailyRate + '/night</p>';
                    html += '<p><strong>Status:</strong> ' + statusText + '</p>';
                    
                    if (!room.available && room.conflicts.length > 0) {
                        html += '<p><strong>Conflict:</strong> ' + room.conflicts[0].guest + ' (' + room.conflicts[0].checkIn + ' - ' + room.conflicts[0].checkOut + ')</p>';
                    }
                    
                    if (room.available) {
                        html += '<button class="btn btn-success" onclick="selectRoomForBooking(\'' + room.roomNumber + '\', \'' + room.dailyRate + '\')">📝 Book This Room</button>';
                    }
                    
                    html += '</div>';
                });
            }
            
            resultsDiv.innerHTML = html;
        }
        
        function selectRoomForBooking(roomNumber, dailyRate) {
            // Switch to booking tab
            showTab('booking');
            document.querySelector('[onclick="showTab(\'booking\')"]').classList.add('active');
            
            // Pre-fill the form
            document.getElementById('booking-room').innerHTML = '<option value="' + roomNumber + '" selected>Room ' + roomNumber + '</option>';
            document.getElementById('daily-rate').value = dailyRate.replace(', '');
            document.getElementById('booking-checkin').value = document.getElementById('avail-checkin').value;
            document.getElementById('booking-checkout').value = document.getElementById('avail-checkout').value;
            
            showStatus('Room ' + roomNumber + ' selected. Please fill in guest details.', 'success');
        }
        
        function createBooking() {
            const guestName = document.getElementById('guest-name').value;
            const checkIn = document.getElementById('booking-checkin').value;
            const checkOut = document.getElementById('booking-checkout').value;
            const roomNumber = document.getElementById('booking-room').value;
            const dailyRate = document.getElementById('daily-rate').value;
            
            if (!guestName || !checkIn || !checkOut || !roomNumber || !dailyRate) {
                showStatus('Please fill in all required fields.', 'error');
                return;
            }
            
            if (new Date(checkIn) >= new Date(checkOut)) {
                showStatus('Check-out date must be after check-in date.', 'error');
                return;
            }
            
            const bookingData = {
                guestName: guestName,
                guestEmail: document.getElementById('guest-email').value,
                guestPhone: document.getElementById('guest-phone').value,
                numberOfGuests: document.getElementById('guest-count').value,
                checkInDate: checkIn,
                checkOutDate: checkOut,
                roomNumber: roomNumber,
                dailyRate: dailyRate,
                purposeOfVisit: document.getElementById('purpose-visit').value,
                specialRequests: document.getElementById('special-requests').value,
                bookingSource: 'Direct',
                paymentStatus: 'Pending'
            };
            
            showStatus('Creating booking...', 'success');
            
            google.script.run
                .withSuccessHandler(function(result) {
                    showStatus(result, 'success');
                    // Clear form
                    document.getElementById('guest-name').value = '';
                    document.getElementById('guest-email').value = '';
                    document.getElementById('guest-phone').value = '';
                    document.getElementById('guest-count').value = '1';
                    document.getElementById('purpose-visit').value = '';
                    document.getElementById('special-requests').value = '';
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, 'error');
                })
                .createNewBooking(JSON.stringify(bookingData));
        }
        
        function processCheckIn() {
            const bookingId = document.getElementById('checkin-booking-id').value;
            
            if (!bookingId) {
                showStatus('Please enter a booking ID.', 'error');
                return;
            }
            
            showStatus('Processing check-in...', 'success');
            
            google.script.run
                .withSuccessHandler(function(result) {
                    showStatus(result, 'success');
                    document.getElementById('checkin-booking-id').value = '';
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, 'error');
                })
                .processCheckIn(JSON.stringify({
                    bookingId: bookingId
                }));
        }
        
        function processCheckOut() {
            const bookingId = document.getElementById('checkout-booking-id').value;
            
            if (!bookingId) {
                showStatus('Please enter a booking ID.', 'error');
                return;
            }
            
            showStatus('Processing check-out...', 'success');
            
            google.script.run
                .withSuccessHandler(function(result) {
                    showStatus(result, 'success');
                    document.getElementById('checkout-booking-id').value = '';
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, 'error');
                })
                .processCheckOut(JSON.stringify({
                    bookingId: bookingId
                }));
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
 * Wrapper functions for menu integration
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
