/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Move-Out Processing Panel - MoveOutPanel.gs
 * 
 * This module provides a user-friendly panel interface to process tenant move-out requests
 * from Google Form responses and update the Tenant sheet accordingly.
 */

const MoveOutPanel = {

  /**
   * Show the move-out processing panel
   */
  showMoveOutProcessingPanel() {
    try {
      console.log('Opening Move-Out Processing Panel...');
      
      const moveOutRequests = this._getMoveOutRequests();
      
      if (moveOutRequests.length === 0) {
        SpreadsheetApp.getUi().alert(
          'No Move-Out Requests Found',
          'No unprocessed move-out requests found. All requests may have been processed already, or no new requests have been submitted.',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        return;
      }
      
      const html = this._generateMoveOutPanelHTML(moveOutRequests);
      const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(900)
        .setHeight(700)
        .setTitle('🏠 Process Move-Out Requests');
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Process Move-Out Requests');
      
    } catch (error) {
      console.error('Error showing move-out panel:', error);
      SpreadsheetApp.getUi().alert('Error', 'Failed to load move-out panel: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  },

  /**
   * Process the selected move-out request and clear tenant from room
   */
  processMoveOutRequest(moveOutData) {
    try {
      console.log('Processing move-out request:', moveOutData);
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const tenantSheet = spreadsheet.getSheetByName(SHEET_NAMES.TENANT);
      const budgetSheet = spreadsheet.getSheetByName(SHEET_NAMES.BUDGET);
      
      if (!tenantSheet) {
        throw new Error('Tenant sheet not found');
      }
      
      // Parse the move-out data
      const data = JSON.parse(moveOutData);
      
      // Find the tenant row
      const tenantRowIndex = this._findTenantRowByNameAndRoom(tenantSheet, data.tenantName, data.roomNumber);
      if (tenantRowIndex === -1) {
        throw new Error(`Tenant ${data.tenantName} in Room ${data.roomNumber} not found in tenant sheet`);
      }
      
      // Get current tenant data before clearing
      const currentTenantData = this._getCurrentTenantData(tenantSheet, tenantRowIndex);
      
      // Clear tenant information from the room
      this._clearTenantFromRoom(tenantSheet, tenantRowIndex);
      
      // Add security deposit return record to budget if applicable
      if (budgetSheet && data.securityDepositReturn && data.securityDepositReturn !== '') {
        this._addSecurityDepositReturn(budgetSheet, data, currentTenantData);
      }
      
      // Mark the move-out request as processed
      this._markMoveOutAsProcessed(data.timestamp, data.tenantName);
      
      console.log(`Move-out processed successfully for ${data.tenantName} from Room ${data.roomNumber}`);
      return `✅ Move-out completed! ${data.tenantName} has been removed from Room ${data.roomNumber} and the room is now available.`;
      
    } catch (error) {
      console.error('Error processing move-out request:', error);
      throw new Error('Failed to process move-out: ' + error.message);
    }
  },

  /**
   * Get move-out requests from form responses
   * @private
   */
  _getMoveOutRequests() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      // Look for the move-out request responses sheet
      let responseSheet = this._findMoveOutRequestSheet(spreadsheet);
      
      if (!responseSheet || responseSheet.getLastRow() <= 1) {
        console.log('No move-out request responses found');
        return [];
      }
      
      const data = responseSheet.getDataRange().getValues();
      const headers = data[0];
      const requests = [];
      
      console.log('Move-out response sheet headers:', headers);
      console.log(`Found ${data.length - 1} total responses in the sheet`);
      
      // Find the Processed column index
      let processedColIndex = headers.findIndex(header => 
        header.toLowerCase().includes('processed')
      );
      
      // Map form headers to our expected fields
      const headerMap = this._createMoveOutHeaderMap(headers);
      
      let processedCount = 0;
      let unprocessedCount = 0;
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // Check if already processed
        let isProcessed = false;
        if (processedColIndex !== -1) {
          const processedValue = row[processedColIndex];
          isProcessed = processedValue && 
                       processedValue.toString().trim() !== '' && 
                       (processedValue.toString().toLowerCase().includes('processed') ||
                        processedValue.toString().toLowerCase().includes('completed'));
        }
        
        if (isProcessed) {
          processedCount++;
          console.log(`Row ${i + 1}: Skipping processed move-out request - ${this._getFieldValue(row, headerMap.tenantName)}`);
          continue;
        }
        
        // Check if tenant name is empty
        const tenantName = this._getFieldValue(row, headerMap.tenantName);
        if (!tenantName || tenantName.toString().trim() === '') {
          console.log(`Row ${i + 1}: Skipping move-out request with empty tenant name`);
          continue;
        }
        
        unprocessedCount++;
        
        const moveOutRequest = {
          timestamp: row[0], // First column is always timestamp
          tenantName: tenantName,
          email: this._getFieldValue(row, headerMap.email) || '',
          phone: this._getFieldValue(row, headerMap.phone) || '',
          roomNumber: this._getFieldValue(row, headerMap.roomNumber) || '',
          plannedMoveOutDate: this._getFieldValue(row, headerMap.plannedMoveOutDate) || '',
          forwardingAddress: this._getFieldValue(row, headerMap.forwardingAddress) || '',
          reasonForMoving: this._getFieldValue(row, headerMap.reasonForMoving) || '',
          additionalDetails: this._getFieldValue(row, headerMap.additionalDetails) || '',
          satisfaction: this._getFieldValue(row, headerMap.satisfaction) || '',
          recommend: this._getFieldValue(row, headerMap.recommend) || '',
          rowIndex: i + 1
        };
        
        requests.push(moveOutRequest);
      }
      
      console.log(`Move-out processing summary: ${processedCount} already processed, ${unprocessedCount} unprocessed requests found`);
      return requests;
      
    } catch (error) {
      console.error('Error getting move-out requests:', error);
      return [];
    }
  },

  /**
   * Find the move-out request sheet by checking multiple possible names
   * @private
   */
  _findMoveOutRequestSheet(spreadsheet) {
    // Try different possible sheet names in order of preference
    const possibleNames = [
      'Move-Out Requests',
      'Move-Out Request',
      'Form Responses 2',
      'Form Responses (2)'
    ];
    
    for (let name of possibleNames) {
      let sheet = spreadsheet.getSheetByName(name);
      if (sheet) {
        console.log(`Found move-out request sheet: ${name}`);
        return sheet;
      }
    }
    
    // If no exact match, look for sheets with move-out headers
    const sheets = spreadsheet.getSheets();
    for (let sheet of sheets) {
      const sheetName = sheet.getName().toLowerCase();
      if (sheetName.includes('form responses') && sheet.getLastColumn() > 0) {
        const headers = sheet.getRange(1, 1, 1, Math.min(10, sheet.getLastColumn())).getValues()[0];
        
        // Check if headers indicate this is a move-out request sheet
        const isMoveOutSheet = headers.some(header => 
          header && (
            header.toString().toLowerCase().includes('tenant name') ||
            header.toString().toLowerCase().includes('planned move-out date') ||
            header.toString().toLowerCase().includes('move-out date') ||
            header.toString().toLowerCase().includes('forwarding address')
          )
        );
        
        if (isMoveOutSheet) {
          console.log(`Found move-out request sheet by headers: ${sheet.getName()}`);
          return sheet;
        }
      }
    }
    
    console.log('Could not find move-out request sheet');
    return null;
  },

  /**
   * Create a mapping between move-out form headers and our expected fields
   * @private
   */
  _createMoveOutHeaderMap(headers) {
    const map = {};
    
    headers.forEach((header, index) => {
      const lowerHeader = header.toLowerCase().trim();
      
      if (lowerHeader === 'tenant name' || lowerHeader.includes('tenant name')) {
        map.tenantName = index;
      } else if (lowerHeader.includes('email')) {
        map.email = index;
      } else if (lowerHeader.includes('phone')) {
        map.phone = index;
      } else if (lowerHeader.includes('room number') || lowerHeader.includes('room')) {
        map.roomNumber = index;
      } else if (lowerHeader.includes('planned move-out date') || lowerHeader.includes('move-out date')) {
        map.plannedMoveOutDate = index;
      } else if (lowerHeader.includes('forwarding address')) {
        map.forwardingAddress = index;
      } else if (lowerHeader.includes('reason for moving') || lowerHeader.includes('primary reason')) {
        map.reasonForMoving = index;
      } else if (lowerHeader.includes('additional details')) {
        map.additionalDetails = index;
      } else if (lowerHeader.includes('satisfaction') || lowerHeader.includes('overall satisfaction')) {
        map.satisfaction = index;
      } else if (lowerHeader.includes('recommend')) {
        map.recommend = index;
      }
    });
    
    console.log('Move-out header mapping created:', map);
    return map;
  },

  /**
   * Find tenant row by name and room number
   * @private
   */
  _findTenantRowByNameAndRoom(sheet, tenantName, roomNumber) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const nameCol = headers.indexOf('Current Tenant Name');
    const roomNumberCol = headers.indexOf('Room Number');
    
    for (let i = 1; i < data.length; i++) {
      const rowTenantName = data[i][nameCol] ? data[i][nameCol].toString().trim() : '';
      const rowRoomNumber = data[i][roomNumberCol] ? data[i][roomNumberCol].toString().trim() : '';
      
      if (rowTenantName === tenantName.trim() && rowRoomNumber === roomNumber.toString().trim()) {
        return i + 1; // Return 1-based row index
      }
    }
    
    return -1; // Not found
  },

  /**
   * Get current tenant data before clearing
   * @private
   */
  _getCurrentTenantData(sheet, rowIndex) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    
    const securityDepositCol = headers.indexOf('Security Deposit Paid');
    
    return {
      securityDeposit: securityDepositCol !== -1 ? rowData[securityDepositCol] : ''
    };
  },

  /**
   * Clear tenant information from room (make it vacant)
   * @private
   */
  _clearTenantFromRoom(sheet, rowIndex) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Find columns to clear
    const nameCol = headers.indexOf('Current Tenant Name');
    const emailCol = headers.indexOf('Tenant Email');
    const phoneCol = headers.indexOf('Tenant Phone');
    const moveInCol = headers.indexOf('Move-In Date');
    const securityDepositCol = headers.indexOf('Security Deposit Paid');
    const statusCol = headers.indexOf('Room Status');
    const lastPaymentCol = headers.indexOf('Last Payment Date');
    const paymentStatusCol = headers.indexOf('Payment Status');
    const emergencyContactCol = headers.indexOf('Emergency Contact');
    const leaseEndCol = headers.indexOf('Lease End Date');
    const notesCol = headers.indexOf('Notes');
    
    // Clear tenant-specific data
    const clearData = [];
    for (let i = 0; i < headers.length; i++) {
      clearData.push(''); // Start with empty values
    }
    
    // Keep room number and rental prices
    const roomNumberCol = headers.indexOf('Room Number');
    const rentalPriceCol = headers.indexOf('Rental Price');
    const negotiatedPriceCol = headers.indexOf('Negotiated Price');
    const moveOutDateCol = headers.indexOf('Move-Out Date (Planned)');
    
    const currentData = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    
    // Preserve room information
    if (roomNumberCol !== -1) clearData[roomNumberCol] = currentData[roomNumberCol];
    if (rentalPriceCol !== -1) clearData[rentalPriceCol] = currentData[rentalPriceCol];
    if (negotiatedPriceCol !== -1) clearData[negotiatedPriceCol] = ''; // Clear negotiated price
    
    // Set room as vacant
    if (statusCol !== -1) clearData[statusCol] = 'Vacant';
    
    // Add move-out note
    if (notesCol !== -1) clearData[notesCol] = `Tenant moved out on ${new Date().toLocaleDateString()}`;
    
    // Set actual move-out date
    if (moveOutDateCol !== -1) clearData[moveOutDateCol] = new Date().toLocaleDateString();
    
    // Update the row
    sheet.getRange(rowIndex, 1, 1, clearData.length).setValues([clearData]);
    
    console.log(`Cleared tenant data from row ${rowIndex}, room is now vacant`);
  },

  /**
   * Add security deposit return record to budget
   * @private
   */
  _addSecurityDepositReturn(budgetSheet, moveOutData, currentTenantData) {
    try {
      const budgetRow = [
        new Date().toLocaleDateString(),     // Date
        'Expense',                           // Type
        `Security Deposit Return - Room ${moveOutData.roomNumber} - ${moveOutData.tenantName}`, // Description
        `-${moveOutData.securityDepositReturn}`, // Amount (negative for expense)
        'Other',                             // Category
        moveOutData.returnMethod || 'Check', // Payment Method
        `DEP-RETURN-${moveOutData.roomNumber}-${new Date().getMonth() + 1}${new Date().getDate()}`, // Reference
        moveOutData.tenantName,              // Tenant/Guest
        ''                                   // Receipt
      ];
      
      const lastRow = budgetSheet.getLastRow();
      budgetSheet.getRange(lastRow + 1, 1, 1, budgetRow.length).setValues([budgetRow]);
      
      console.log(`Added security deposit return to budget: ${moveOutData.securityDepositReturn} for ${moveOutData.tenantName}`);
      
    } catch (error) {
      console.error('Error adding security deposit return to budget:', error);
    }
  },

  /**
   * Mark move-out request as processed
   * @private
   */
  _markMoveOutAsProcessed(timestamp, tenantName) {
    try {
      console.log(`Marking move-out request as processed for: ${tenantName}`);
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const moveOutSheet = this._findMoveOutRequestSheet(spreadsheet);
      
      if (!moveOutSheet) {
        console.log('Could not find move-out request sheet to mark as processed');
        return;
      }
      
      const data = moveOutSheet.getDataRange().getValues();
      const headers = data[0];
      
      // Find or create Processed column
      let processedCol = headers.findIndex(header => 
        header && header.toString().toLowerCase().includes('processed')
      );
      
      if (processedCol === -1) {
        console.log('Adding Processed column to move-out sheet...');
        processedCol = headers.length;
        moveOutSheet.getRange(1, processedCol + 1).setValue('Processed');
        
        // Format the header
        moveOutSheet.getRange(1, processedCol + 1)
          .setBackground('#1c4587')
          .setFontColor('white')
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
        
        moveOutSheet.setColumnWidth(processedCol + 1, 150);
      }
      
      // Find columns for name matching
      const nameCol = headers.findIndex(header => 
        header && header.toString().toLowerCase().includes('tenant name')
      );
      
      // Find and mark the specific request
      let found = false;
      for (let i = 1; i < data.length; i++) {
        const rowName = data[i][nameCol] ? data[i][nameCol].toString().trim() : '';
        const rowTimestamp = data[i][0];
        
        // Match by both timestamp and name for accuracy
        if (rowTimestamp.toString() === timestamp.toString() && rowName === tenantName.trim()) {
          const processedValue = `Completed - ${new Date().toLocaleDateString()}`;
          moveOutSheet.getRange(i + 1, processedCol + 1).setValue(processedValue);
          
          // Add conditional formatting
          const cell = moveOutSheet.getRange(i + 1, processedCol + 1);
          cell.setBackground('#d9ead3'); // Light green
          
          found = true;
          console.log(`✅ Successfully marked move-out request as processed: ${tenantName}`);
          break;
        }
      }
      
      if (!found) {
        console.log(`❌ Could not find move-out request for: ${tenantName} with timestamp: ${timestamp}`);
      }
      
    } catch (error) {
      console.error('Error marking move-out request as processed:', error);
    }
  },

  /**
   * Get field value safely
   * @private
   */
  _getFieldValue(row, index) {
    return (index !== undefined && row[index] !== undefined) ? row[index] : '';
  },

  /**
   * Generate HTML for the move-out processing panel
   * @private
   */
  _generateMoveOutPanelHTML(moveOutRequests) {
    const requestsJson = JSON.stringify(moveOutRequests);
    
    return `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background: #dc3545; color: white; padding: 15px; margin: -20px -20px 20px -20px; }
        .selector-section {
            background: #f8f9fa;
            border: 1px solid #ddd;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 20px;
        }
        .request-card { 
            border: 1px solid #ddd; 
            border-radius: 8px; 
            margin: 15px 0; 
            padding: 15px; 
            background: #f9f9f9; 
            display: none;
        }
        .request-header { 
            background: #f8d7da; 
            padding: 10px; 
            margin: -15px -15px 15px -15px; 
            border-radius: 8px 8px 0 0; 
            font-weight: bold; 
        }
        .request-details { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin: 10px 0; }
        .request-field { margin: 5px 0; }
        .request-field strong { color: #dc3545; }
        .processing-section { 
            background: #fff; 
            border: 1px solid #ccc; 
            border-radius: 5px; 
            padding: 20px; 
            margin-top: 15px; 
        }
        .form-group { margin: 20px 0; }
        .form-group label { display: block; font-weight: bold; margin-bottom: 10px; color: #333; }
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
        .form-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-bottom: 20px;
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
        .btn-secondary { 
            background: #6c757d; 
            margin-left: 10px;
        }
        .btn-secondary:hover { background: #5a6268; }
        .status { padding: 10px; margin: 10px 0; border-radius: 4px; }
        .status.success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .status.error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .no-requests { text-align: center; color: #666; margin: 50px 0; }
        .instruction { color: #666; font-style: italic; margin-bottom: 20px; text-align: center; }
        .feedback-section {
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            border-radius: 5px;
            padding: 15px;
            margin: 15px 0;
        }
        .feedback-title { font-weight: bold; color: #856404; margin-bottom: 10px; }
    </style>
</head>
<body>
    <div class="header">
        <h2>🏠 Process Move-Out Requests</h2>
        <p>Select a tenant to process their move-out request and clear them from the system</p>
    </div>
    
    <div id="status" class="status" style="display: none;"></div>
    
    ${moveOutRequests.length === 0 ? `
        <div class="no-requests">
            <h3>No Pending Move-Out Requests</h3>
            <p>All move-out requests have been processed or no new requests have been submitted.</p>
        </div>
    ` : `
        <div class="selector-section">
            <h3>Select Move-Out Request to Process</h3>
            <div class="instruction">Choose a tenant from the dropdown below to view their move-out request and process their departure.</div>
            <div class="form-group">
                <label for="request-selector">Tenant Move-Out Request:</label>
                <select id="request-selector" onchange="showSelectedRequest()">
                    <option value="">-- Choose a move-out request --</option>
                    ${moveOutRequests.map((request, index) => `
                        <option value="${index}">${request.tenantName} - Room ${request.roomNumber} (Move-Out: ${new Date(request.plannedMoveOutDate).toLocaleDateString()})</option>
                    `).join('')}
                </select>
            </div>
        </div>
        
        ${moveOutRequests.map((request, index) => `
            <div class="request-card" id="request-${index}">
                <div class="request-header">
                    Move-Out Request from ${request.tenantName}
                    <span style="float: right; font-size: 12px;">Submitted: ${new Date(request.timestamp).toLocaleDateString()}</span>
                </div>
                
                <div class="request-details">
                    <div>
                        <div class="request-field"><strong>Tenant Name:</strong> ${request.tenantName}</div>
                        <div class="request-field"><strong>Room Number:</strong> ${request.roomNumber}</div>
                        <div class="request-field"><strong>Email:</strong> ${request.email}</div>
                        <div class="request-field"><strong>Phone:</strong> ${request.phone}</div>
                        <div class="request-field"><strong>Planned Move-Out Date:</strong> ${new Date(request.plannedMoveOutDate).toLocaleDateString()}</div>
                    </div>
                    <div>
                        <div class="request-field"><strong>Reason for Moving:</strong> ${request.reasonForMoving || 'Not specified'}</div>
                        <div class="request-field"><strong>Overall Satisfaction:</strong> ${request.satisfaction || 'Not provided'}</div>
                        <div class="request-field"><strong>Would Recommend:</strong> ${request.recommend || 'Not provided'}</div>
                        <div class="request-field"><strong>Forwarding Address:</strong> ${request.forwardingAddress || 'Not provided'}</div>
                    </div>
                </div>
                
                ${request.additionalDetails ? `
                    <div class="request-field" style="margin-top: 15px;">
                        <strong>Additional Details:</strong><br>
                        <em>${request.additionalDetails}</em>
                    </div>
                ` : ''}
                
                <div class="feedback-section">
                    <div class="feedback-title">📝 Tenant Feedback Summary</div>
                    <div><strong>Satisfaction Level:</strong> ${request.satisfaction || 'Not provided'}</div>
                    <div><strong>Would Recommend:</strong> ${request.recommend || 'Not provided'}</div>
                    <div><strong>Reason for Leaving:</strong> ${request.reasonForMoving || 'Not specified'}</div>
                </div>
                
                <div class="processing-section">
                    <h4 style="margin-bottom: 25px; color: #dc3545;">🏠 Process Move-Out</h4>
                    
                    <div class="form-row">
                        <div class="form-group">
                            <label>Actual Move-Out Date:</label>
                            <input type="date" id="moveout-date-${index}" required>
                            <small>The actual date the tenant moved out (defaults to today)</small>
                        </div>
                        <div class="form-group">
                            <label>Security Deposit Return:</label>
                            <input type="text" id="deposit-return-${index}" placeholder="e.g., $1200">
                            <small>Amount of security deposit to return (leave blank if none)</small>
                        </div>
                    </div>
                    
                    <div class="form-row">
                        <div class="form-group">
                            <label>Return Method:</label>
                            <select id="return-method-${index}">
                                <option value="Check">Check</option>
                                <option value="Bank Transfer">Bank Transfer</option>
                                <option value="Cash">Cash</option>
                                <option value="Zelle">Zelle</option>
                                <option value="Other">Other</option>
                            </select>
                        </div>
                        <div class="form-group">
                            <label>Final Notes:</label>
                            <input type="text" id="notes-${index}" placeholder="Optional notes about the move-out">
                        </div>
                    </div>
                    
                    <div style="margin-top: 30px; text-align: center;">
                        <button class="btn" onclick="confirmMoveOut(${index})">🏠 Process Move-Out</button>
                        <button class="btn btn-secondary" onclick="clearMoveOutForm(${index})">🔄 Clear Form</button>
                    </div>
                </div>
            </div>
        `).join('')}
    `}
    
    <script>
        const moveOutRequests = ${requestsJson};
        
        // Set today's date as default move-out date
        document.addEventListener('DOMContentLoaded', function() {
            const today = new Date().toISOString().split('T')[0];
            moveOutRequests.forEach((request, index) => {
                const dateField = document.getElementById('moveout-date-' + index);
                if (dateField) {
                    dateField.value = today;
                }
            });
        });
        
        function showSelectedRequest() {
            const selectedIndex = document.getElementById('request-selector').value;
            
            // Hide all request cards
            moveOutRequests.forEach((request, index) => {
                document.getElementById('request-' + index).style.display = 'none';
            });
            
            // Show selected request card
            if (selectedIndex !== '') {
                document.getElementById('request-' + selectedIndex).style.display = 'block';
            }
        }
        
        function confirmMoveOut(index) {
            const request = moveOutRequests[index];
            const moveOutDate = document.getElementById('moveout-date-' + index).value;
            const securityDepositReturn = document.getElementById('deposit-return-' + index).value;
            const returnMethod = document.getElementById('return-method-' + index).value;
            const notes = document.getElementById('notes-' + index).value;
            
            if (!moveOutDate) {
                showStatus('Please select the actual move-out date.', 'error');
                return;
            }
            
            const confirmMessage = 'Are you sure you want to process the move-out for ' + request.tenantName + ' from Room ' + request.roomNumber + '? This will remove them from the tenant sheet and mark the room as vacant.';
            
            if (!confirm(confirmMessage)) {
                return;
            }
            
            const moveOutData = {
                timestamp: request.timestamp,
                tenantName: request.tenantName,
                email: request.email,
                phone: request.phone,
                roomNumber: request.roomNumber,
                plannedMoveOutDate: request.plannedMoveOutDate,
                forwardingAddress: request.forwardingAddress,
                reasonForMoving: request.reasonForMoving,
                additionalDetails: request.additionalDetails,
                satisfaction: request.satisfaction,
                recommend: request.recommend,
                actualMoveOutDate: moveOutDate,
                securityDepositReturn: securityDepositReturn,
                returnMethod: returnMethod,
                finalNotes: notes
            };
            
            showStatus('Processing move-out...', 'success');
            
            google.script.run
                .withSuccessHandler(function(result) {
                    showStatus(result, 'success');
                    clearMoveOutForm(index);
                    // Remove the processed request from the dropdown
                    const option = document.querySelector('#request-selector option[value="' + index + '"]');
                    if (option) option.remove();
                    // Hide the card
                    document.getElementById('request-' + index).style.display = 'none';
                    document.getElementById('request-selector').value = '';
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, 'error');
                })
                .processMoveOutRequest(JSON.stringify(moveOutData));
        }
        
        function clearMoveOutForm(index) {
            document.getElementById('moveout-date-' + index).value = new Date().toISOString().split('T')[0];
            document.getElementById('deposit-return-' + index).value = '';
            document.getElementById('return-method-' + index).value = 'Check';
            document.getElementById('notes-' + index).value = '';
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
function showMoveOutProcessingPanel() {
  return MoveOutPanel.showMoveOutProcessingPanel();
}

/**
 * Server-side function called from the HTML panel
 */
function processMoveOutRequest(moveOutData) {
  return MoveOutPanel.processMoveOutRequest(moveOutData);
}
