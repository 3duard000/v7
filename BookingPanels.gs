/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Individual Booking Panels - BookingPanels.gs
 * 
 * Separate panels for each booking function
 */

const BookingPanels = {

  /**
   * Show availability checker panel
   */
  showAvailabilityChecker() {
    try {
      const html = this._generateAvailabilityHTML();
      const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(800)
        .setHeight(600)
        .setTitle('⚡ Check Room Availability');
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Check Room Availability');
    } catch (error) {
      console.error('Error showing availability checker:', error);
      SpreadsheetApp.getUi().alert('Error', 'Failed to load availability checker: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  },

  /**
   * Show new booking panel
   */
  showNewBookingPanel() {
    try {
      const html = this._generateNewBookingHTML();
      const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(800)
        .setHeight(700)
        .setTitle('📝 Create New Booking');
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Create New Booking');
    } catch (error) {
      console.error('Error showing new booking panel:', error);
      SpreadsheetApp.getUi().alert('Error', 'Failed to load new booking panel: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  },

  /**
   * Show check-in panel
   */
  showCheckInPanel() {
    try {
      const html = this._generateCheckInHTML();
      const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(700)
        .setHeight(500)
        .setTitle('✅ Process Check-In');
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Process Check-In');
    } catch (error) {
      console.error('Error showing check-in panel:', error);
      SpreadsheetApp.getUi().alert('Error', 'Failed to load check-in panel: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  },

  /**
   * Show online reservation processing panel
   */
  showOnlineReservationPanel() {
    try {
      const html = this._generateOnlineReservationHTML();
      const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(900)
        .setHeight(700)
        .setTitle('🏨 Process Online Reservations');
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Process Online Reservations');
    } catch (error) {
      console.error('Error showing online reservation panel:', error);
      SpreadsheetApp.getUi().alert('Error', 'Failed to load online reservation panel: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  },

  /**
   * Process online reservation from Google Form
   */
  processOnlineReservation(reservationData) {
    try {
      console.log('Processing online reservation:', reservationData);
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const guestSheet = spreadsheet.getSheetByName(SHEET_NAMES.GUEST_ROOMS);
      
      if (!guestSheet) {
        throw new Error('Guest Rooms sheet not found');
      }
      
      const data = JSON.parse(reservationData);
      
      // Generate booking ID and calculate totals
      const bookingId = this._generateBookingId();
      const nights = this._calculateNights(new Date(data.checkInDate), new Date(data.checkOutDate));
      const totalAmount = parseFloat(data.dailyRate) * nights;
      
      // Create booking row
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
        'Confirmed',                         // Payment Status
        'Confirmed',                         // Booking Status
        'Google Form',                       // Source
        `Online reservation processed on ${new Date().toLocaleDateString()}` // Notes
      ];
      
      // Add to guest rooms sheet
      const lastRow = guestSheet.getLastRow();
      guestSheet.getRange(lastRow + 1, 1, 1, bookingRow.length).setValues([bookingRow]);
      
      // Add to Google Calendar
      BookingManager._addToGoogleCalendar({
        type: 'guest',
        title: `Guest: ${data.guestName}`,
        startDate: data.checkInDate,
        endDate: data.checkOutDate,
        room: data.roomNumber,
        details: `Online reservation - Room ${data.roomNumber}\nGuest: ${data.guestName}\nEmail: ${data.guestEmail}\nPhone: ${data.guestPhone}`
      });
      
      console.log(`Online reservation ${bookingId} processed successfully`);
      return `✅ Online reservation ${bookingId} confirmed for ${data.guestName} in Room ${data.roomNumber}`;
      
    } catch (error) {
      console.error('Error processing online reservation:', error);
      throw new Error('Failed to process online reservation: ' + error.message);
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
   * Generate availability checker HTML
   * @private
   */
  _generateAvailabilityHTML() {
    return `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background: #1c4587; color: white; padding: 15px; margin: -20px -20px 20px -20px; }
        .form-group { margin: 20px 0; }
        .form-group label { display: block; font-weight: bold; margin-bottom: 10px; }
        .form-group input { 
            width: calc(100% - 24px); 
            padding: 12px; 
            border: 1px solid #ccc; 
            border-radius: 4px; 
            font-size: 14px;
            box-sizing: border-box;
        }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }
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
        .status { padding: 10px; margin: 10px 0; border-radius: 4px; }
        .status.success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .status.error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .room-card { 
            border: 1px solid #ddd; 
            border-radius: 8px; 
            padding: 15px; 
            margin: 10px 0; 
            background: #f9f9f9;
        }
        .room-available { border-left: 4px solid #22803c; }
        .room-occupied { border-left: 4px solid #dc3545; }
        .availability-results { margin-top: 20px; }
    </style>
</head>
<body>
    <div class="header">
        <h2>⚡ Check Room Availability</h2>
        <p>Enter dates to see which rooms are available</p>
    </div>
    
    <div id="status" class="status" style="display: none;"></div>
    
    <div class="form-row">
        <div class="form-group">
            <label>Check-In Date:</label>
            <input type="date" id="checkin-date" required>
        </div>
        <div class="form-group">
            <label>Check-Out Date:</label>
            <input type="date" id="checkout-date" required>
        </div>
    </div>
    
    <button class="btn" onclick="checkAvailability()">🔍 Check Availability</button>
    
    <div id="availability-results" class="availability-results"></div>
    
    <script>
        // Set default dates
        document.addEventListener('DOMContentLoaded', function() {
            const today = new Date();
            const tomorrow = new Date(today);
            tomorrow.setDate(tomorrow.getDate() + 1);
            
            document.getElementById('checkin-date').value = today.toISOString().split('T')[0];
            document.getElementById('checkout-date').value = tomorrow.toISOString().split('T')[0];
        });
        
        function checkAvailability() {
            const checkIn = document.getElementById('checkin-date').value;
            const checkOut = document.getElementById('checkout-date').value;
            
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
                        const dailyRateNum = parseFloat(room.dailyRate.replace('$', '')) || 75;
                        const totalCost = dailyRateNum * data.nights;
                        html += '<p><strong>Total Cost:</strong> $' + totalCost.toFixed(2) + ' for ' + data.nights + ' nights</p>';
                    }
                    
                    html += '</div>';
                });
            }
            
            resultsDiv.innerHTML = html;
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
</html>`;
  },

  /**
   * Generate new booking HTML
   * @private
   */
  _generateNewBookingHTML() {
    return `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background: #22803c; color: white; padding: 15px; margin: -20px -20px 20px -20px; }
        .form-group { margin: 20px 0; }
        .form-group label { display: block; font-weight: bold; margin-bottom: 10px; }
        .form-group input, .form-group select, .form-group textarea { 
            width: calc(100% - 24px); 
            padding: 12px; 
            border: 1px solid #ccc; 
            border-radius: 4px; 
            font-size: 14px;
            box-sizing: border-box;
        }
        .form-group textarea { resize: vertical; min-height: 80px; }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }
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
        .status { padding: 10px; margin: 10px 0; border-radius: 4px; }
        .status.success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .status.error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .required { color: #dc3545; }
    </style>
</head>
<body>
    <div class="header">
        <h2>📝 Create New Booking</h2>
        <p>Fill in guest details to create a new booking</p>
    </div>
    
    <div id="status" class="status" style="display: none;"></div>
    
    <div class="form-row">
        <div class="form-group">
            <label>Guest Name <span class="required">*</span>:</label>
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
            <label>Check-In Date <span class="required">*</span>:</label>
            <input type="date" id="checkin-date" required>
        </div>
        <div class="form-group">
            <label>Check-Out Date <span class="required">*</span>:</label>
            <input type="date" id="checkout-date" required>
        </div>
    </div>
    
    <div class="form-row">
        <div class="form-group">
            <label>Room Number <span class="required">*</span>:</label>
            <input type="text" id="room-number" placeholder="e.g., 201" required>
        </div>
        <div class="form-group">
            <label>Daily Rate <span class="required">*</span>:</label>
            <input type="number" id="daily-rate" step="0.01" min="0" placeholder="75.00" required>
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
    
    <button class="btn" onclick="createBooking()">📝 Create Booking</button>
    
    <script>
        // Set default dates
        document.addEventListener('DOMContentLoaded', function() {
            const today = new Date();
            const tomorrow = new Date(today);
            tomorrow.setDate(tomorrow.getDate() + 1);
            
            document.getElementById('checkin-date').value = today.toISOString().split('T')[0];
            document.getElementById('checkout-date').value = tomorrow.toISOString().split('T')[0];
        });
        
        function createBooking() {
            const guestName = document.getElementById('guest-name').value;
            const checkIn = document.getElementById('checkin-date').value;
            const checkOut = document.getElementById('checkout-date').value;
            const roomNumber = document.getElementById('room-number').value;
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
                    document.getElementById('room-number').value = '';
                    document.getElementById('daily-rate').value = '';
                    document.getElementById('purpose-visit').value = '';
                    document.getElementById('special-requests').value = '';
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, 'error');
                })
                .createNewBooking(JSON.stringify(bookingData));
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
</html>`;
  },

  /**
   * Generate check-in HTML
   * @private
   */
  _generateCheckInHTML() {
    return `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background: #28a745; color: white; padding: 15px; margin: -20px -20px 20px -20px; }
        .form-group { margin: 20px 0; }
        .form-group label { display: block; font-weight: bold; margin-bottom: 10px; }
        .form-group input { 
            width: calc(100% - 24px); 
            padding: 12px; 
            border: 1px solid #ccc; 
            border-radius: 4px; 
            font-size: 14px;
            box-sizing: border-box;
        }
        .btn { 
            background: #28a745; 
            color: white; 
            padding: 12px 25px; 
            border: none; 
            border-radius: 4px; 
            cursor: pointer; 
            font-size: 14px;
            font-weight: bold;
        }
        .btn:hover { background: #218838; }
        .status { padding: 10px; margin: 10px 0; border-radius: 4px; }
        .status.success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .status.error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .info-box { background: #e7f3ff; border: 1px solid #b3d7ff; padding: 15px; border-radius: 8px; margin: 20px 0; }
    </style>
</head>
<body>
    <div class="header">
        <h2>✅ Process Check-In</h2>
        <p>Enter booking ID to check in a guest</p>
    </div>
    
    <div id="status" class="status" style="display: none;"></div>
    
    <div class="info-box">
        <h4>Check-In Process</h4>
        <p>Enter the guest's booking ID to process their check-in. This will update their status to "Occupied" and record the check-in date.</p>
    </div>
    
    <div class="form-group">
        <label>Booking ID:</label>
        <input type="text" id="booking-id" placeholder="e.g., BK123456" required>
    </div>
    
    <button class="btn" onclick="processCheckIn()">✅ Process Check-In</button>
    
    <script>
        function processCheckIn() {
            const bookingId = document.getElementById('booking-id').value;
            
            if (!bookingId) {
                showStatus('Please enter a booking ID.', 'error');
                return;
            }
            
            showStatus('Processing check-in...', 'success');
            
            google.script.run
                .withSuccessHandler(function(result) {
                    showStatus(result, 'success');
                    document.getElementById('booking-id').value = '';
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, 'error');
                })
                .processCheckIn(JSON.stringify({
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
</html>`;
  },

  /**
   * Generate online reservation processing HTML
   * @private
   */
  _generateOnlineReservationHTML() {
    return `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background: #6f42c1; color: white; padding: 15px; margin: -20px -20px 20px -20px; }
        .form-group { margin: 20px 0; }
        .form-group label { display: block; font-weight: bold; margin-bottom: 10px; }
        .form-group input, .form-group select { 
            width: calc(100% - 24px); 
            padding: 12px; 
            border: 1px solid #ccc; 
            border-radius: 4px; 
            font-size: 14px;
            box-sizing: border-box;
        }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }
        .btn { 
            background: #6f42c1; 
            color: white; 
            padding: 12px 25px; 
            border: none; 
            border-radius: 4px; 
            cursor: pointer; 
            font-size: 14px;
            font-weight: bold;
        }
        .btn:hover { background: #5a32a3; }
        .status { padding: 10px; margin: 10px 0; border-radius: 4px; }
        .status.success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .status.error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .info-box { background: #e7e3ff; border: 1px solid #c7b3ff; padding: 15px; border-radius: 8px; margin: 20px 0; }
        .required { color: #dc3545; }
    </style>
</head>
<body>
    <div class="header">
        <h2>🏨 Process Online Reservations</h2>
        <p>Review and confirm guest reservations from Google Forms</p>
    </div>
    
    <div id="status" class="status" style="display: none;"></div>
    
    <div class="info-box">
        <h4>Online Reservation Processing</h4>
        <p>Use this panel to review Google Form submissions and convert them into confirmed bookings. Fill in the guest details and room assignment below.</p>
    </div>
    
    <div class="form-row">
        <div class="form-group">
            <label>Guest Name <span class="required">*</span>:</label>
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
            <label>Check-In Date <span class="required">*</span>:</label>
            <input type="date" id="checkin-date" required>
        </div>
        <div class="form-group">
            <label>Check-Out Date <span class="required">*</span>:</label>
            <input type="date" id="checkout-date" required>
        </div>
    </div>
    
    <div class="form-row">
        <div class="form-group">
            <label>Assign Room <span class="required">*</span>:</label>
            <input type="text" id="room-number" placeholder="e.g., 201" required>
        </div>
        <div class="form-group">
            <label>Daily Rate <span class="required">*</span>:</label>
            <input type="number" id="daily-rate" step="0.01" min="0" placeholder="75.00" required>
        </div>
    </div>
    
    <div class="form-group">
        <label>Purpose of Visit:</label>
        <input type="text" id="purpose-visit" placeholder="Business, vacation, etc.">
    </div>
    
    <button class="btn" onclick="processReservation()">🏨 Confirm Online Reservation</button>
    
    <script>
        // Set default dates
        document.addEventListener('DOMContentLoaded', function() {
            const today = new Date();
            const tomorrow = new Date(today);
            tomorrow.setDate(tomorrow.getDate() + 1);
            
            document.getElementById('checkin-date').value = today.toISOString().split('T')[0];
            document.getElementById('checkout-date').value = tomorrow.toISOString().split('T')[0];
        });
        
        function processReservation() {
            const guestName = document.getElementById('guest-name').value;
            const checkIn = document.getElementById('checkin-date').value;
            const checkOut = document.getElementById('checkout-date').value;
            const roomNumber = document.getElementById('room-number').value;
            const dailyRate = document.getElementById('daily-rate').value;
            
            if (!guestName || !checkIn || !checkOut || !roomNumber || !dailyRate) {
                showStatus('Please fill in all required fields.', 'error');
                return;
            }
            
            if (new Date(checkIn) >= new Date(checkOut)) {
                showStatus('Check-out date must be after check-in date.', 'error');
                return;
            }
            
            const reservationData = {
                guestName: guestName,
                guestEmail: document.getElementById('guest-email').value,
                guestPhone: document.getElementById('guest-phone').value,
                numberOfGuests: document.getElementById('guest-count').value,
                checkInDate: checkIn,
                checkOutDate: checkOut,
                roomNumber: roomNumber,
                dailyRate: dailyRate,
                purposeOfVisit: document.getElementById('purpose-visit').value
            };
            
            showStatus('Processing online reservation...', 'success');
            
            google.script.run
                .withSuccessHandler(function(result) {
                    showStatus(result, 'success');
                    // Clear form
                    document.getElementById('guest-name').value = '';
                    document.getElementById('guest-email').value = '';
                    document.getElementById('guest-phone').value = '';
                    document.getElementById('guest-count').value = '1';
                    document.getElementById('room-number').value = '';
                    document.getElementById('daily-rate').value = '';
                    document.getElementById('purpose-visit').value = '';
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, 'error');
                })
                .processOnlineReservation(JSON.stringify(reservationData));
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
</html>`;
  },
  _generateCheckOutHTML() {
    return `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background: #dc3545; color: white; padding: 15px; margin: -20px -20px 20px -20px; }
        .form-group { margin: 20px 0; }
        .form-group label { display: block; font-weight: bold; margin-bottom: 10px; }
        .form-group input { 
            width: calc(100% - 24px); 
            padding: 12px; 
            border: 1px solid #ccc; 
            border-radius: 4px; 
            font-size: 14px;
            box-sizing: border-box;
        }
        .btn { 
            background: #dc3545; 
            color: white; 
            padding: 12px 25px; 
            border: none; 
            border-radius: 4px; 
            cursor: pointer; 
            font-size: 14px;
            font-weight: bold;
        }
        .btn:hover { background: #c82333; }
        .status { padding: 10px; margin: 10px 0; border-radius: 4px; }
        .status.success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .status.error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .info-box { background: #fff3cd; border: 1px solid #ffeaa7; padding: 15px; border-radius: 8px; margin: 20px 0; }
    </style>
</head>
<body>
    <div class="header">
        <h2>🚪 Process Check-Out</h2>
        <p>Enter booking ID to check out a guest</p>
    </div>
    
    <div id="status" class="status" style="display: none;"></div>
    
    <div class="info-box">
        <h4>Check-Out Process</h4>
        <p>Enter the guest's booking ID to process their check-out. This will clear the guest information and mark the room as available.</p>
    </div>
    
    <div class="form-group">
        <label>Booking ID:</label>
        <input type="text" id="booking-id" placeholder="e.g., BK123456" required>
    </div>
    
    <button class="btn" onclick="processCheckOut()">🚪 Process Check-Out</button>
    
    <script>
        function processCheckOut() {
            const bookingId = document.getElementById('booking-id').value;
            
            if (!bookingId) {
                showStatus('Please enter a booking ID.', 'error');
                return;
            }
            
            showStatus('Processing check-out...', 'success');
            
            google.script.run
                .withSuccessHandler(function(result) {
                    showStatus(result, 'success');
                    document.getElementById('booking-id').value = '';
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
</html>`;
  }
};

/**
 * Wrapper functions for menu integration
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
