/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Maintenance Request Entry Panel - MaintenancePanel.gs
 * 
 * This module provides a user-friendly panel interface to add maintenance requests
 * directly to the Maintenance Requests sheet.
 */

const MaintenancePanel = {

  /**
   * Show the maintenance request entry panel
   */
  showMaintenanceRequestPanel() {
    try {
      console.log('Opening Maintenance Request Panel...');
      
      const html = this._generatePanelHTML();
      const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(800)
        .setHeight(650)
        .setTitle('🔧 Add Maintenance Request');
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Add Maintenance Request');
      
    } catch (error) {
      console.error('Error showing maintenance request panel:', error);
      SpreadsheetApp.getUi().alert('Error', 'Failed to load maintenance request panel: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  },

  /**
   * Add new maintenance request to the sheet
   */
  addMaintenanceRequest(requestData) {
    try {
      console.log('Adding maintenance request:', requestData);
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const maintenanceSheet = spreadsheet.getSheetByName(SHEET_NAMES.MAINTENANCE);
      
      if (!maintenanceSheet) {
        throw new Error('Maintenance Requests sheet not found');
      }
      
      // Parse the request data
      const data = JSON.parse(requestData);
      
      // Generate unique request ID
      const requestId = this._generateRequestId();
      
      // Create maintenance request row data
      const maintenanceRow = [
        requestId,                           // Request ID
        new Date().toLocaleString(),         // Timestamp
        data.roomArea,                       // Room/Area
        data.issueType,                      // Issue Type
        data.priority,                       // Priority
        data.description,                    // Description
        data.reportedBy || 'Property Manager', // Reported By
        data.contactInfo || '',              // Contact Info
        data.assignedTo || '',               // Assigned To
        'Pending',                           // Status
        data.estimatedCost || '',            // Estimated Cost
        '',                                  // Actual Cost (empty for new request)
        '',                                  // Date Started (empty for new request)
        '',                                  // Date Completed (empty for new request)
        data.partsNeeded || '',              // Parts Used
        data.estimatedHours || '',           // Labor Hours
        '',                                  // Photos (empty for new request)
        data.notes || `Created on ${new Date().toLocaleDateString()}` // Notes
      ];
      
      // Add the new maintenance request to the sheet
      const lastRow = maintenanceSheet.getLastRow();
      maintenanceSheet.getRange(lastRow + 1, 1, 1, maintenanceRow.length).setValues([maintenanceRow]);
      
      console.log(`Maintenance request ${requestId} added successfully`);
      return `✅ Maintenance request ${requestId} has been created successfully for ${data.roomArea}.`;
      
    } catch (error) {
      console.error('Error adding maintenance request:', error);
      throw new Error('Failed to add maintenance request: ' + error.message);
    }
  },

  /**
   * Get all rooms/areas from tenant and guest sheets for dropdown
   * @private
   */
  _getRoomsAndAreas() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const rooms = new Set();
      
      // Add tenant rooms
      const tenantSheet = spreadsheet.getSheetByName(SHEET_NAMES.TENANT);
      if (tenantSheet && tenantSheet.getLastRow() > 1) {
        const tenantData = tenantSheet.getDataRange().getValues();
        const headers = tenantData[0];
        const roomNumberCol = headers.indexOf('Room Number');
        
        for (let i = 1; i < tenantData.length; i++) {
          if (tenantData[i][roomNumberCol]) {
            rooms.add(`Room ${tenantData[i][roomNumberCol]}`);
          }
        }
      }
      
      // Add guest rooms
      const guestSheet = spreadsheet.getSheetByName(SHEET_NAMES.GUEST_ROOMS);
      if (guestSheet && guestSheet.getLastRow() > 1) {
        const guestData = guestSheet.getDataRange().getValues();
        const headers = guestData[0];
        const roomNumberCol = headers.indexOf('Room Number');
        const roomNameCol = headers.indexOf('Room Name');
        
        for (let i = 1; i < guestData.length; i++) {
          if (guestData[i][roomNumberCol]) {
            const roomName = guestData[i][roomNameCol] || 'Guest Room';
            rooms.add(`Room ${guestData[i][roomNumberCol]} - ${roomName}`);
          }
        }
      }
      
      // Add common areas
      const commonAreas = [
        'Common Area - Lobby',
        'Common Area - Kitchen',
        'Common Area - Dining Room',
        'Common Area - Living Room',
        'Common Area - Bathroom',
        'Common Area - Hallway',
        'Common Area - Stairs',
        'Common Area - Laundry Room',
        'Exterior - Front Yard',
        'Exterior - Back Yard',
        'Exterior - Parking Area',
        'Exterior - Entrance',
        'Building Systems - HVAC',
        'Building Systems - Plumbing',
        'Building Systems - Electrical',
        'Building Systems - Internet/WiFi'
      ];
      
      commonAreas.forEach(area => rooms.add(area));
      
      return Array.from(rooms).sort();
    } catch (error) {
      console.error('Error getting rooms and areas:', error);
      return ['Room 101', 'Room 102', 'Common Area'];
    }
  },

  /**
   * Get maintenance staff/contractors for assignment dropdown
   * @private
   */
  _getMaintenanceStaff() {
    // This could be expanded to read from a configuration sheet
    return [
      '',
      'Mike Johnson - General Maintenance',
      'Tom Rodriguez - HVAC Specialist',
      'Sarah Williams - Plumbing',
      'Paint Crew - Painting & Repairs',
      'Cleaning Team - Deep Cleaning',
      'External Contractor - TBD'
    ];
  },

  /**
   * Generate a unique request ID
   * @private
   */
  _generateRequestId() {
    const prefix = 'MR';
    const timestamp = Date.now().toString().slice(-6);
    return `${prefix}-${timestamp}`;
  },

  /**
   * Generate HTML for the maintenance request panel
   * @private
   */
  _generatePanelHTML() {
    const roomsAndAreas = this._getRoomsAndAreas();
    const maintenanceStaff = this._getMaintenanceStaff();
    
    return `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background: #ff6d00; color: white; padding: 15px; margin: -20px -20px 20px -20px; }
        .form-container {
            background: #f8f9fa;
            border: 1px solid #ddd;
            border-radius: 8px;
            padding: 25px;
            margin-bottom: 20px;
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
            color: #ff6d00;
            margin-bottom: 20px;
            padding-bottom: 8px;
            border-bottom: 2px solid #ff6d00;
        }
        .form-group { margin: 20px 0; }
        .form-group label { display: block; font-weight: bold; margin-bottom: 10px; }
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
            gap: 25px;
            margin-bottom: 25px;
        }
        .form-row-three {
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            gap: 20px;
            margin-bottom: 25px;
        }
        .btn { 
            background: #ff6d00; 
            color: white; 
            padding: 12px 25px; 
            border: none; 
            border-radius: 4px; 
            cursor: pointer; 
            font-size: 14px;
            font-weight: bold;
        }
        .btn:hover { background: #e65100; }
        .btn-secondary { 
            background: #6c757d; 
            margin-left: 10px;
        }
        .btn-secondary:hover { background: #5a6268; }
        .status { padding: 10px; margin: 10px 0; border-radius: 4px; }
        .status.success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .status.error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .required { color: #dc3545; }
        .priority-high { background-color: #ffebee; }
        .priority-emergency { background-color: #ffcdd2; }
    </style>
</head>
<body>
    <div class="header">
        <h2>🔧 Add Maintenance Request</h2>
        <p>Create a new maintenance request for the White House property</p>
    </div>
    
    <div id="status" class="status" style="display: none;"></div>
    
    <div class="form-container">
        <form id="maintenanceForm">
            <div class="form-section">
                <div class="section-title">Request Details</div>
                <div class="form-row">
                    <div class="form-group">
                        <label>Room/Area <span class="required">*</span>:</label>
                        <select id="roomArea" required>
                            <option value="">-- Select Room/Area --</option>
                            ${roomsAndAreas.map(room => `<option value="${room}">${room}</option>`).join('')}
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Issue Type <span class="required">*</span>:</label>
                        <select id="issueType" required>
                            <option value="">-- Select Issue Type --</option>
                            ${DROPDOWN_OPTIONS.MAINTENANCE.ISSUE_TYPE.map(type => `<option value="${type}">${type}</option>`).join('')}
                        </select>
                    </div>
                </div>
                
                <div class="form-group">
                    <label>Description <span class="required">*</span>:</label>
                    <textarea id="description" placeholder="Describe the maintenance issue in detail..." required></textarea>
                    <small>Provide as much detail as possible to help maintenance staff understand the issue</small>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>Priority <span class="required">*</span>:</label>
                        <select id="priority" onchange="updatePriorityStyle()" required>
                            <option value="">-- Select Priority --</option>
                            ${DROPDOWN_OPTIONS.MAINTENANCE.PRIORITY.map(priority => `<option value="${priority}">${priority}</option>`).join('')}
                        </select>
                        <small>Emergency: Immediate safety risk | High: Affects habitability | Medium: Should be fixed soon | Low: Can wait</small>
                    </div>
                    <div class="form-group">
                        <label>Estimated Cost:</label>
                        <input type="text" id="estimatedCost" placeholder="e.g., $150">
                        <small>Optional - your best estimate for materials and labor</small>
                    </div>
                </div>
            </div>
            
            <div class="form-section">
                <div class="section-title">Contact Information</div>
                <div class="form-row">
                    <div class="form-group">
                        <label>Reported By:</label>
                        <input type="text" id="reportedBy" placeholder="Your name" value="Property Manager">
                    </div>
                    <div class="form-group">
                        <label>Contact Info:</label>
                        <input type="text" id="contactInfo" placeholder="Email or phone number">
                        <small>How maintenance staff can reach you for questions</small>
                    </div>
                </div>
            </div>
            
            <div class="form-section">
                <div class="section-title">Assignment & Planning</div>
                <div class="form-row">
                    <div class="form-group">
                        <label>Assign To:</label>
                        <select id="assignedTo">
                            ${maintenanceStaff.map(staff => `<option value="${staff}">${staff || '-- Not Assigned --'}</option>`).join('')}
                        </select>
                        <small>Leave blank to assign later</small>
                    </div>
                    <div class="form-group">
                        <label>Estimated Hours:</label>
                        <input type="number" id="estimatedHours" placeholder="e.g., 2" step="0.5" min="0">
                        <small>Estimated time to complete the work</small>
                    </div>
                </div>
                
                <div class="form-group">
                    <label>Parts/Materials Needed:</label>
                    <input type="text" id="partsNeeded" placeholder="e.g., Pipe fitting, LED bulb, Paint">
                    <small>List any parts or materials that may be needed</small>
                </div>
                
                <div class="form-group">
                    <label>Additional Notes:</label>
                    <textarea id="notes" placeholder="Any additional information, special instructions, or context..."></textarea>
                </div>
            </div>
            
            <div style="margin-top: 30px; text-align: center;">
                <button type="button" class="btn" onclick="submitMaintenanceRequest()">🔧 Create Request</button>
                <button type="button" class="btn btn-secondary" onclick="resetForm()">🔄 Clear Form</button>
            </div>
        </form>
    </div>
    
    <script>
        function updatePriorityStyle() {
            const priority = document.getElementById('priority').value;
            const select = document.getElementById('priority');
            
            // Remove existing priority classes
            select.classList.remove('priority-high', 'priority-emergency');
            
            // Add appropriate class based on priority
            if (priority === 'High') {
                select.classList.add('priority-high');
            } else if (priority === 'Emergency') {
                select.classList.add('priority-emergency');
            }
        }
        
        function submitMaintenanceRequest() {
            // Validate required fields
            const roomArea = document.getElementById('roomArea').value;
            const issueType = document.getElementById('issueType').value;
            const description = document.getElementById('description').value;
            const priority = document.getElementById('priority').value;
            
            if (!roomArea || !issueType || !description || !priority) {
                showStatus('Please fill in all required fields (marked with *).', 'error');
                return;
            }
            
            const requestData = {
                roomArea: roomArea,
                issueType: issueType,
                description: description,
                priority: priority,
                estimatedCost: document.getElementById('estimatedCost').value,
                reportedBy: document.getElementById('reportedBy').value,
                contactInfo: document.getElementById('contactInfo').value,
                assignedTo: document.getElementById('assignedTo').value,
                estimatedHours: document.getElementById('estimatedHours').value,
                partsNeeded: document.getElementById('partsNeeded').value,
                notes: document.getElementById('notes').value
            };
            
            showStatus('Creating maintenance request...', 'success');
            
            google.script.run
                .withSuccessHandler(function(result) {
                    showStatus(result, 'success');
                    resetForm();
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, 'error');
                })
                .addMaintenanceRequest(JSON.stringify(requestData));
        }
        
        function resetForm() {
            document.getElementById('maintenanceForm').reset();
            document.getElementById('reportedBy').value = 'Property Manager';
            document.getElementById('priority').classList.remove('priority-high', 'priority-emergency');
            showStatus('Form cleared. Ready for new request.', 'success');
        }
        
        function showStatus(message, type) {
            const status = document.getElementById('status');
            status.textContent = message;
            status.className = 'status ' + type;
            status.style.display = 'block';
            
            setTimeout(() => {
                if (type === 'success' && message.includes('created successfully')) {
                    status.style.display = 'none';
                }
            }, 5000);
        }
        
        // Auto-focus on first field when form loads
        document.addEventListener('DOMContentLoaded', function() {
            document.getElementById('roomArea').focus();
        });
    </script>
</body>
</html>
    `;
  }
};

/**
 * Wrapper function for menu integration
 */
function showMaintenanceRequestPanel() {
  return MaintenancePanel.showMaintenanceRequestPanel();
}

/**
 * Server-side function called from the HTML panel
 */
function addMaintenanceRequest(requestData) {
  return MaintenancePanel.addMaintenanceRequest(requestData);
}
