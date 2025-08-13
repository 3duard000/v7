/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Application Processing Panel - TenantPanel.gs
 * 
 * This module provides a user-friendly panel interface to process tenant applications
 * from Google Form responses and add approved applicants to the Tenant sheet.
 * 
 * UPDATES:
 * - Fixed processed applications filtering to prevent re-showing processed entries
 * - Fixed Monthly Income display to show actual form values
 * - FIXED: Proper status updates when approving/rejecting applications
 */

const Panel = {

  /**
   * Show the tenant application processing panel
   */
  showApplicationProcessingPanel() {
    try {
      console.log('Opening Application Processing Panel...');
      
      const applications = this._getTenantApplications();
      
      if (applications.length === 0) {
        SpreadsheetApp.getUi().alert(
          'No Applications Found',
          'No unprocessed tenant applications found. All applications may have been processed already, or no new applications have been submitted.',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        return;
      }
      
      const html = this._generatePanelHTML(applications);
      const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(900)
        .setHeight(700)
        .setTitle('🏠 Process Tenant Applications');
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Process Tenant Applications');
      
    } catch (error) {
      console.error('Error showing application panel:', error);
      SpreadsheetApp.getUi().alert('Error', 'Failed to load application panel: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  },

  /**
   * Process the selected application and update existing tenant row
   */
  processApplication(applicationData) {
    try {
      console.log('Processing application:', applicationData);
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const tenantSheet = spreadsheet.getSheetByName(SHEET_NAMES.TENANT);
      
      if (!tenantSheet) {
        throw new Error('Tenant sheet not found');
      }
      
      // Parse the application data
      const data = JSON.parse(applicationData);
      
      // Find the row with the selected room number
      const roomRowIndex = this._findRoomRow(tenantSheet, data.roomNumber);
      if (roomRowIndex === -1) {
        throw new Error(`Room ${data.roomNumber} not found in tenant sheet`);
      }
      
      // Update the existing row with tenant information
      const tenantRow = [
        data.roomNumber,           // Room Number (keep existing)
        data.rentAmount,           // Rental Price (keep existing)
        data.negotiatedRent,       // Negotiated Price
        data.fullName,             // Current Tenant Name
        data.email,                // Tenant Email
        data.phone,                // Tenant Phone
        data.moveInDate,           // Move-In Date
        data.securityDeposit,      // Security Deposit Paid
        'Occupied',                // Room Status
        '',                        // Last Payment Date (empty for new tenant)
        'Current',                 // Payment Status (default to current)
        data.leaseEndDate,         // Move-Out Date (Planned)
        data.emergencyContact,     // Emergency Contact
        data.leaseEndDate,         // Lease End Date
        `Approved from application on ${new Date().toLocaleDateString()}`  // Notes
      ];
      
      // Update the existing row instead of adding new one
      tenantSheet.getRange(roomRowIndex, 1, 1, tenantRow.length).setValues([tenantRow]);
      
      // Mark the application as approved in the responses sheet
      this._markApplicationAsProcessed(data.timestamp, data.fullName, 'Approved', data.email);
      
      console.log(`Application processed successfully for ${data.fullName} in Room ${data.roomNumber}`);
      return `✅ Application approved! ${data.fullName} has been assigned to Room ${data.roomNumber}.`;
      
    } catch (error) {
      console.error('Error processing application:', error);
      throw new Error('Failed to process application: ' + error.message);
    }
  },

  /**
   * Mark application as rejected (called from HTML)
   */
  markApplicationAsRejected(applicationData) {
    try {
      const data = JSON.parse(applicationData);
      console.log('Rejecting application for:', data.fullName);
      
      // Mark as rejected in the form responses sheet
      this._markApplicationAsProcessed(data.timestamp, data.fullName, 'Rejected', data.email);
      
      return `✅ Application for ${data.fullName} has been rejected and marked as processed.`;
      
    } catch (error) {
      console.error('Error rejecting application:', error);
      throw new Error('Failed to reject application: ' + error.message);
    }
  },

  /**
   * Get tenant applications from form responses - IMPROVED FILTERING
   * @private
   */
  _getTenantApplications() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      // Look for the tenant application responses sheet
      let responseSheet = this._findTenantApplicationSheet(spreadsheet);
      
      if (!responseSheet || responseSheet.getLastRow() <= 1) {
        console.log('No tenant application responses found');
        return [];
      }
      
      const data = responseSheet.getDataRange().getValues();
      const headers = data[0];
      const applications = [];
      
      console.log('Response sheet headers:', headers);
      console.log(`Found ${data.length - 1} total responses in the sheet`);
      
      // Find the Processed column index
      let processedColIndex = headers.findIndex(header => 
        header.toLowerCase().includes('processed')
      );
      
      // Map form headers to our expected fields
      const headerMap = this._createHeaderMap(headers);
      
      let processedCount = 0;
      let unprocessedCount = 0;
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // IMPROVED: Check if already processed more thoroughly
        let isProcessed = false;
        if (processedColIndex !== -1) {
          const processedValue = row[processedColIndex];
          isProcessed = processedValue && 
                       processedValue.toString().trim() !== '' && 
                       processedValue.toString().trim().toLowerCase() !== 'pending review' &&
                       (processedValue.toString().toLowerCase().includes('approved') ||
                        processedValue.toString().toLowerCase().includes('rejected'));
        }
        
        if (isProcessed) {
          processedCount++;
          console.log(`Row ${i + 1}: Skipping processed application - ${this._getFieldValue(row, headerMap.fullName)} (Status: ${row[processedColIndex]})`);
          continue; // Skip processed applications
        }
        
        // Also check if the applicant name is empty (incomplete submission)
        const fullName = this._getFieldValue(row, headerMap.fullName);
        if (!fullName || fullName.toString().trim() === '') {
          console.log(`Row ${i + 1}: Skipping application with empty name`);
          continue;
        }
        
        unprocessedCount++;
        
        const application = {
          timestamp: row[0], // First column is always timestamp
          fullName: fullName,
          email: this._getFieldValue(row, headerMap.email) || '',
          phone: this._getFieldValue(row, headerMap.phone) || '',
          currentAddress: this._getFieldValue(row, headerMap.currentAddress) || '',
          moveInDate: this._getFieldValue(row, headerMap.moveInDate) || '',
          preferredRoom: this._getFieldValue(row, headerMap.preferredRoom) || '',
          employment: this._getFieldValue(row, headerMap.employment) || '',
          employer: this._getFieldValue(row, headerMap.employer) || '',
          monthlyIncome: this._getFieldValue(row, headerMap.monthlyIncome) || '',
          reference1: this._getFieldValue(row, headerMap.reference1) || '',
          reference2: this._getFieldValue(row, headerMap.reference2) || '',
          emergencyContact: this._getFieldValue(row, headerMap.emergencyContact) || '',
          aboutYourself: this._getFieldValue(row, headerMap.aboutYourself) || '',
          specialNeeds: this._getFieldValue(row, headerMap.specialNeeds) || '',
          proofOfIncome: this._getFieldValue(row, headerMap.proofOfIncome) || '',
          rowIndex: i + 1 // Store row index for processing
        };
        
        // Debug logging for Monthly Income
        console.log(`Application ${application.fullName}: Monthly Income = "${application.monthlyIncome}"`);
        
        applications.push(application);
      }
      
      console.log(`Processing summary: ${processedCount} already processed, ${unprocessedCount} unprocessed applications found`);
      return applications;
      
    } catch (error) {
      console.error('Error getting tenant applications:', error);
      return [];
    }
  },

  /**
   * Find the tenant application sheet by checking multiple possible names
   * @private
   */
  _findTenantApplicationSheet(spreadsheet) {
    // Try different possible sheet names in order of preference
    const possibleNames = [
      'Tenant Application',
      'Tenant Applications', 
      'Form Responses 1',
      'Form Responses (1)'
    ];
    
    for (let name of possibleNames) {
      let sheet = spreadsheet.getSheetByName(name);
      if (sheet) {
        console.log(`Found tenant application sheet: ${name}`);
        return sheet;
      }
    }
    
    // If no exact match, look for sheets with tenant application headers
    const sheets = spreadsheet.getSheets();
    for (let sheet of sheets) {
      const sheetName = sheet.getName().toLowerCase();
      if (sheetName.includes('form responses') && sheet.getLastColumn() > 0) {
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
          console.log(`Found tenant application sheet by headers: ${sheet.getName()}`);
          return sheet;
        }
      }
    }
    
    console.log('Could not find tenant application sheet');
    return null;
  },

  /**
   * Create a mapping between form headers and our expected fields - IMPROVED
   * @private
   */
  _createHeaderMap(headers) {
    const map = {};
    
    headers.forEach((header, index) => {
      const lowerHeader = header.toLowerCase().trim();
      
      // Be more specific about Full Name - prioritize exact matches
      if (lowerHeader === 'full name' || lowerHeader.startsWith('full name')) {
        map.fullName = index;
      } else if (lowerHeader === 'name' && !map.fullName) {
        // Only use 'name' if we haven't found 'full name' yet
        map.fullName = index;
      } else if (lowerHeader.includes('email')) {
        map.email = index;
      } else if (lowerHeader.includes('phone')) {
        map.phone = index;
      } else if (lowerHeader.includes('current address') || lowerHeader.includes('address')) {
        map.currentAddress = index;
      } else if (lowerHeader.includes('move-in date') || lowerHeader.includes('move in') || lowerHeader.includes('desired move-in date')) {
        map.moveInDate = index;
      } else if (lowerHeader.includes('preferred room') || lowerHeader.includes('room')) {
        map.preferredRoom = index;
      } else if (lowerHeader.includes('employment status') || lowerHeader.includes('employment')) {
        map.employment = index;
      } else if (lowerHeader.includes('employer') || lowerHeader.includes('school')) {
        map.employer = index;
      } else if (lowerHeader === 'monthly income') {
        map.monthlyIncome = index; // FIXED: Only exact match for "monthly income"
      } else if (lowerHeader.includes('proof of income') || lowerHeader.includes('proof')) {
        map.proofOfIncome = index; // Separate mapping for proof of income
      } else if (lowerHeader.includes('reference 1') || lowerHeader.includes('reference1')) {
        map.reference1 = index;
      } else if (lowerHeader.includes('reference 2') || lowerHeader.includes('reference2')) {
        map.reference2 = index;
      } else if (lowerHeader.includes('emergency contact')) {
        map.emergencyContact = index;
      } else if (lowerHeader.includes('tell us about yourself') || lowerHeader.includes('about yourself')) {
        map.aboutYourself = index;
      } else if (lowerHeader.includes('special needs') || lowerHeader.includes('requests')) {
        map.specialNeeds = index;
      }
    });
    
    // Debug logging to help identify any mapping issues
    console.log('Header mapping created:', map);
    console.log('Available headers:', headers);
    console.log('Monthly Income mapped to index:', map.monthlyIncome);
    if (map.monthlyIncome !== undefined) {
      console.log('Monthly Income header:', headers[map.monthlyIncome]);
    }
    
    return map;
  },

  /**
   * Get available rooms from tenant sheet
   * @private
   */
  _getAvailableRooms() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const tenantSheet = spreadsheet.getSheetByName(SHEET_NAMES.TENANT);
      
      if (!tenantSheet || tenantSheet.getLastRow() <= 1) {
        return [];
      }
      
      const data = tenantSheet.getDataRange().getValues();
      const headers = data[0];
      const rooms = [];
      
      const roomNumberCol = headers.indexOf('Room Number');
      const rentalPriceCol = headers.indexOf('Rental Price');
      const statusCol = headers.indexOf('Room Status');
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[roomNumberCol]) {
          const roomNumber = row[roomNumberCol].toString();
          const rentalPrice = row[rentalPriceCol] || '';
          const status = row[statusCol] || 'Available';
          
          let statusDisplay = '';
          switch (status.toLowerCase()) {
            case 'occupied':
              statusDisplay = '(Occupied)';
              break;
            case 'maintenance':
              statusDisplay = '(Maintenance)';
              break;
            case 'reserved':
              statusDisplay = '(Reserved)';
              break;
            case 'vacant':
            default:
              statusDisplay = '(Available)';
              break;
          }
          
          rooms.push({
            number: roomNumber,
            price: rentalPrice,
            status: status,
            display: `${roomNumber} ${statusDisplay}`,
            available: status.toLowerCase() === 'vacant' || status.toLowerCase() === 'available'
          });
        }
      }
      
      return rooms;
    } catch (error) {
      console.error('Error getting available rooms:', error);
      return [];
    }
  },

  /**
   * Find the row index for a specific room number
   * @private
   */
  _findRoomRow(sheet, roomNumber) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const roomNumberCol = headers.indexOf('Room Number');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][roomNumberCol].toString() === roomNumber.toString()) {
        return i + 1; // Return 1-based row index
      }
    }
    
    return -1; // Room not found
  },

  /**
   * Get field value safely
   * @private
   */
  _getFieldValue(row, index) {
    return (index !== undefined && row[index] !== undefined) ? row[index] : '';
  },

  /**
   * Generate HTML for the application processing panel
   * @private
   */
  _generatePanelHTML(applications) {
    const applicationsJson = JSON.stringify(applications);
    const availableRooms = this._getAvailableRooms();
    const roomsJson = JSON.stringify(availableRooms);
    
    return `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background: #1c4587; color: white; padding: 15px; margin: -20px -20px 20px -20px; }
        .selector-section {
            background: #f8f9fa;
            border: 1px solid #ddd;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 20px;
        }
        .application-card { 
            border: 1px solid #ddd; 
            border-radius: 8px; 
            margin: 15px 0; 
            padding: 15px; 
            background: #f9f9f9; 
            display: none;
        }
        .app-header { 
            background: #e8f0fe; 
            padding: 10px; 
            margin: -15px -15px 15px -15px; 
            border-radius: 8px 8px 0 0; 
            font-weight: bold; 
        }
        .app-details { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin: 10px 0; }
        .app-field { margin: 5px 0; }
        .app-field strong { color: #1c4587; }
        .approval-section { 
            background: #fff; 
            border: 1px solid #ccc; 
            border-radius: 5px; 
            padding: 15px; 
            margin-top: 15px; 
        }
        .form-group { margin: 20px 0; }
        .form-group label { display: block; font-weight: bold; margin-bottom: 10px; }
        .form-group input, .form-group select { 
            width: 100%; 
            padding: 12px; 
            border: 1px solid #ccc; 
            border-radius: 4px; 
            font-size: 14px;
            margin-bottom: 8px;
        }
        .form-group small {
            color: #666; 
            font-size: 11px; 
            margin-top: 8px; 
            display: block;
            line-height: 1.4;
        }
        .btn { 
            background: #22803c; 
            color: white; 
            padding: 10px 20px; 
            border: none; 
            border-radius: 4px; 
            cursor: pointer; 
            font-size: 14px; 
        }
        .btn:hover { background: #1a6b30; }
        .btn-reject { background: #cc0000; }
        .btn-reject:hover { background: #990000; }
        .status { padding: 10px; margin: 10px 0; border-radius: 4px; }
        .status.success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .status.error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .no-applications { text-align: center; color: #666; margin: 50px 0; }
        .occupied { color: #cc0000; }
        .maintenance { color: #ff6d00; }
        .available { color: #22803c; }
        .instruction { color: #666; font-style: italic; margin-bottom: 20px; text-align: center; }
        .income-highlight { 
            background: #fff3cd; 
            border: 1px solid #ffeaa7; 
            padding: 8px; 
            border-radius: 4px; 
            font-weight: bold;
        }
    </style>
</head>
<body>
    <div class="header">
        <h2>🏠 Process Tenant Applications</h2>
        <p>Select an applicant to review and process their application</p>
    </div>
    
    <div id="status" class="status" style="display: none;"></div>
    
    ${applications.length === 0 ? `
        <div class="no-applications">
            <h3>No Pending Applications</h3>
            <p>All applications have been processed or no new applications have been submitted.</p>
        </div>
    ` : `
        <div class="selector-section">
            <h3>Select Applicant to Process</h3>
            <div class="instruction">Choose an applicant from the dropdown below to view their application details and process approval.</div>
            <div class="form-group">
                <label for="applicant-selector">Applicant Name:</label>
                <select id="applicant-selector" onchange="showSelectedApplication()">
                    <option value="">-- Choose an applicant --</option>
                    ${applications.map((app, index) => `
                        <option value="${index}">${app.fullName} (Applied: ${new Date(app.timestamp).toLocaleDateString()})</option>
                    `).join('')}
                </select>
            </div>
        </div>
        
        ${applications.map((app, index) => `
            <div class="application-card" id="app-${index}">
                <div class="app-header">
                    Application from ${app.fullName}
                    <span style="float: right; font-size: 12px;">Submitted: ${new Date(app.timestamp).toLocaleDateString()}</span>
                </div>
                
                <div class="app-details">
                    <div>
                        <div class="app-field"><strong>Full Name:</strong> ${app.fullName}</div>
                        <div class="app-field"><strong>Email:</strong> ${app.email}</div>
                        <div class="app-field"><strong>Phone:</strong> ${app.phone}</div>
                        <div class="app-field"><strong>Current Address:</strong> ${app.currentAddress}</div>
                        <div class="app-field"><strong>Desired Move-in Date:</strong> ${app.moveInDate}</div>
                        <div class="app-field"><strong>Preferred Room:</strong> ${app.preferredRoom || 'Any available'}</div>
                    </div>
                    <div>
                        <div class="app-field"><strong>Employment Status:</strong> ${app.employment}</div>
                        <div class="app-field"><strong>Employer/School:</strong> ${app.employer}</div>
                        <div class="app-field income-highlight"><strong>Monthly Income:</strong> ${app.monthlyIncome}</div>
                        <div class="app-field"><strong>Emergency Contact:</strong> ${app.emergencyContact}</div>
                        <div class="app-field"><strong>Reference 1:</strong> ${app.reference1}</div>
                        ${app.reference2 ? `<div class="app-field"><strong>Reference 2:</strong> ${app.reference2}</div>` : ''}
                    </div>
                </div>
                
                <div class="app-field">
                    <strong>About the Applicant:</strong><br>
                    <em>${app.aboutYourself || 'No additional information provided.'}</em>
                </div>
                
                ${app.specialNeeds ? `
                    <div class="app-field">
                        <strong>Special Requests/Needs:</strong><br>
                        <em>${app.specialNeeds}</em>
                    </div>
                ` : ''}
                
                <div class="approval-section">
                    <h4 style="margin-bottom: 25px;">Approval Decision</h4>
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 25px; margin-bottom: 25px;">
                        <div class="form-group">
                            <label>Assign Room:</label>
                            <select id="room-${index}" onchange="updateRentForApplication(${index})">
                                <option value="">Choose a room...</option>
                                ${availableRooms.map(room => `
                                    <option value="${room.number}" data-price="${room.price}" class="${room.status.toLowerCase()}">${room.display} - ${room.price}</option>
                                `).join('')}
                            </select>
                        </div>
                        <div class="form-group">
                            <label>Security Deposit:</label>
                            <input type="text" id="deposit-${index}" placeholder="e.g., $1200" required>
                        </div>
                    </div>
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 25px; margin-bottom: 25px;">
                        <div class="form-group">
                            <label>Lease End Date:</label>
                            <input type="date" id="lease-end-${index}" required>
                        </div>
                        <div class="form-group">
                            <label>Negotiated Rent:</label>
                            <input type="text" id="negotiated-${index}" placeholder="Optional - leave blank if same as room rent">
                            <small>Only fill if different from the room's standard rent</small>
                        </div>
                    </div>
                    <div style="margin-top: 30px;">
                        <button class="btn" onclick="approveApplication(${index})">✅ Approve Application</button>
                        <button class="btn btn-reject" onclick="rejectApplication(${index})" style="margin-left: 10px;">❌ Reject Application</button>
                    </div>
                </div>
            </div>
        `).join('')}
    `}
    
    <script>
        const applications = ${applicationsJson};
        const availableRooms = ${roomsJson};
        
        function showSelectedApplication() {
            const selectedIndex = document.getElementById('applicant-selector').value;
            
            // Hide all application cards
            applications.forEach((app, index) => {
                document.getElementById('app-' + index).style.display = 'none';
            });
            
            // Show selected application card
            if (selectedIndex !== '') {
                const selectedApp = applications[selectedIndex];
                document.getElementById('app-' + selectedIndex).style.display = 'block';
                
                // Auto-populate lease end date (1 year from move-in date)
                if (selectedApp.moveInDate) {
                    const moveInDate = new Date(selectedApp.moveInDate);
                    const leaseEndDate = new Date(moveInDate);
                    leaseEndDate.setFullYear(leaseEndDate.getFullYear() + 1);
                    document.getElementById('lease-end-' + selectedIndex).value = leaseEndDate.toISOString().split('T')[0];
                }
                
                // Debug log for Monthly Income
                console.log('Selected application Monthly Income:', selectedApp.monthlyIncome);
            }
        }
        
        function updateRentForApplication(index) {
            const roomSelect = document.getElementById('room-' + index);
            const selectedOption = roomSelect.options[roomSelect.selectedIndex];
            if (selectedOption && selectedOption.dataset.price) {
                // Auto-populate security deposit with room price
                document.getElementById('deposit-' + index).value = selectedOption.dataset.price;
            }
        }
        
        function approveApplication(index) {
            const app = applications[index];
            const roomNumber = document.getElementById('room-' + index).value;
            const securityDeposit = document.getElementById('deposit-' + index).value;
            const leaseEndDate = document.getElementById('lease-end-' + index).value;
            const negotiatedRent = document.getElementById('negotiated-' + index).value;
            
            if (!roomNumber || !securityDeposit || !leaseEndDate) {
                showStatus('Please fill in all required fields (Room, Security Deposit, and Lease End Date).', 'error');
                return;
            }
            
            const selectedRoom = availableRooms.find(room => room.number === roomNumber);
            
            if (!selectedRoom) {
                showStatus('Selected room not found.', 'error');
                return;
            }
            
            const applicationData = {
                ...app,
                roomNumber: roomNumber,
                rentAmount: selectedRoom.price,
                negotiatedRent: negotiatedRent || selectedRoom.price,
                securityDeposit: securityDeposit,
                leaseEndDate: leaseEndDate
            };
            
            showStatus('Processing application...', 'success');
            
            google.script.run
                .withSuccessHandler(function(result) {
                    showStatus(result, 'success');
                    // Reset the form
                    document.getElementById('applicant-selector').value = '';
                    document.getElementById('app-' + index).style.display = 'none';
                    // Remove the processed application from the dropdown
                    const option = document.querySelector('#applicant-selector option[value="' + index + '"]');
                    if (option) option.remove();
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, 'error');
                })
                .processApplication(JSON.stringify(applicationData));
        }
        
        function rejectApplication(index) {
            if (confirm('Are you sure you want to reject this application? This action cannot be undone.')) {
                const app = applications[index];
                
                showStatus('Rejecting application...', 'error');
                
                // Mark as rejected in the backend
                google.script.run
                    .withSuccessHandler(function(result) {
                        showStatus(result, 'error');
                        // Reset the form and hide the application
                        document.getElementById('applicant-selector').value = '';
                        document.getElementById('app-' + index).style.display = 'none';
                        // Remove the rejected application from the dropdown
                        const option = document.querySelector('#applicant-selector option[value="' + index + '"]');
                        if (option) option.remove();
                    })
                    .withFailureHandler(function(error) {
                        showStatus('Error marking as rejected: ' + error.message, 'error');
                    })
                    .markApplicationAsRejected(JSON.stringify({
                        timestamp: app.timestamp,
                        fullName: app.fullName,
                        email: app.email
                    }));
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
  },

  /**
   * Mark application as processed in the responses sheet - FIXED VERSION WITH NAME+EMAIL MATCHING
   * @private
   */
  _markApplicationAsProcessed(timestamp, applicantName, status = 'Approved', applicantEmail = '') {
    try {
      console.log(`Marking application as ${status} for: ${applicantName} (${applicantEmail})`);
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      // Find the tenant application sheet using our improved method
      const tenantSheet = this._findTenantApplicationSheet(spreadsheet);
      
      if (!tenantSheet) {
        console.log('Could not find tenant application sheet to mark as processed');
        return;
      }
      
      console.log(`Found sheet: ${tenantSheet.getName()}`);
      
      const data = tenantSheet.getDataRange().getValues();
      const headers = data[0];
      
      // Find the Processed column
      let processedCol = headers.findIndex(header => 
        header && header.toString().toLowerCase().includes('processed')
      );
      
      if (processedCol === -1) {
        console.log('Could not find Processed column, adding it...');
        // Add Processed column if it doesn't exist
        processedCol = headers.length;
        tenantSheet.getRange(1, processedCol + 1).setValue('Processed');
        
        // Format the header
        tenantSheet.getRange(1, processedCol + 1)
          .setBackground('#1c4587')
          .setFontColor('white')
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
        
        // Add dropdown validation
        const dropdownOptions = ['Pending Review', 'Approved', 'Rejected'];
        const validationRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(dropdownOptions)
          .setAllowInvalid(false)
          .setHelpText('Select the processing status for this application')
          .build();
        
        const dropdownRange = tenantSheet.getRange(2, processedCol + 1, 500, 1);
        dropdownRange.setDataValidation(validationRule);
        tenantSheet.setColumnWidth(processedCol + 1, 150);
      }
      
      // Find columns for name and email matching
      const nameCol = headers.findIndex(header => 
        header && header.toString().toLowerCase().includes('full name')
      );
      const emailCol = headers.findIndex(header => 
        header && header.toString().toLowerCase().includes('email')
      );
      
      console.log(`Looking for name in column ${nameCol + 1}, email in column ${emailCol + 1}`);
      
      // Find and mark the specific application using name + email combination
      let found = false;
      for (let i = 1; i < data.length; i++) {
        const rowName = data[i][nameCol] ? data[i][nameCol].toString().trim() : '';
        const rowEmail = data[i][emailCol] ? data[i][emailCol].toString().trim().toLowerCase() : '';
        const targetName = applicantName.toString().trim();
        const targetEmail = applicantEmail.toString().trim().toLowerCase();
        
        console.log(`Row ${i + 1}: Comparing "${rowName}" vs "${targetName}" AND "${rowEmail}" vs "${targetEmail}"`);
        
        // Match by both name and email for accuracy
        if (rowName === targetName && rowEmail === targetEmail) {
          console.log(`✅ Found matching application at row ${i + 1}, setting status to: ${status}`);
          tenantSheet.getRange(i + 1, processedCol + 1).setValue(status);
          found = true;
          
          // Also add conditional formatting based on status
          const cell = tenantSheet.getRange(i + 1, processedCol + 1);
          if (status === 'Approved') {
            cell.setBackground('#d9ead3'); // Light green
          } else if (status === 'Rejected') {
            cell.setBackground('#f4cccc'); // Light red
          }
          
          break;
        }
      }
      
      if (found) {
        console.log(`✅ Successfully marked application as ${status}: ${applicantName}`);
      } else {
        console.log(`❌ Could not find application for: ${applicantName} (${applicantEmail})`);
        // Log available applications for debugging
        console.log('Available applications in sheet:');
        for (let i = 1; i < Math.min(data.length, 6); i++) {
          const rowName = data[i][nameCol] || 'No Name';
          const rowEmail = data[i][emailCol] || 'No Email';
          console.log(`  Row ${i + 1}: "${rowName}" - "${rowEmail}"`);
        }
      }
      
    } catch (error) {
      console.error('Error marking application as processed:', error);
    }
  }
};

/**
 * Wrapper function for menu integration
 */
function showApplicationProcessingPanel() {
  return Panel.showApplicationProcessingPanel();
}

/**
 * Server-side function called from the HTML panel
 */
function processApplication(applicationData) {
  return Panel.processApplication(applicationData);
}

/**
 * Server-side function called from the HTML panel to mark applications as rejected
 */
function markApplicationAsRejected(applicationData) {
  return Panel.markApplicationAsRejected(applicationData);
}
