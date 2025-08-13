/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Data Manager - DataManager.gs
 * 
 * This module handles sample data creation with proper expense formatting (negative amounts).
 */

const DataManager = {

  /**
   * Add sample data to all sheets for demonstration purposes
   */
  addSampleData() {
    console.log('Adding sample data to sheets...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Add sample data to each sheet
    this._addTenantSampleData(spreadsheet);
    this._addGuestRoomsSampleData(spreadsheet);
    this._addBudgetSampleData(spreadsheet);
    this._addMaintenanceSampleData(spreadsheet);
    
    console.log('Sample data addition completed!');
  },

  /**
   * Add sample data to Tenant sheet
   * @private
   */
  _addTenantSampleData(spreadsheet) {
    const tenantSheet = spreadsheet.getSheetByName(SHEET_NAMES.TENANT);
    if (tenantSheet && tenantSheet.getLastRow() === 1) { // Only add if no data exists
      const tenantData = [
        ['101', '$1200', '$1150', 'John Smith', 'john.smith@email.com', '(555) 123-4567', '2024-01-15', '$1150', 'Occupied', '2024-07-01', 'Current', '2025-01-15', 'Jane Smith - (555) 123-4568', '2024-12-31', 'Quiet tenant, pays on time'],
        ['102', '$1300', '$1250', 'Maria Garcia', 'maria.garcia@email.com', '(555) 234-5678', '2024-03-01', '$1250', 'Occupied', '2024-06-15', 'Late', '2025-03-01', 'Carlos Garcia - (555) 234-5679', '2025-02-28', 'Student, sometimes late on payments'],
        ['103', '$1400', '$1400', 'David Johnson', 'david.johnson@email.com', '(555) 345-6789', '2024-05-01', '$1400', 'Occupied', '2024-08-01', 'Current', '', 'Susan Johnson - (555) 345-6790', '2025-04-30', 'Works from home, very clean'],
        ['104', '$1100', '$1050', '', '', '', '', '', 'Vacant', '', '', '', '', '', 'Available for new tenant'],
        ['105', '$1200', '', '', '', '', '', '', 'Maintenance', '', '', '', '', '', 'Painting and carpet cleaning needed']
      ];
      tenantSheet.getRange(2, 1, tenantData.length, tenantData[0].length).setValues(tenantData);
      console.log('✅ Sample data added to Tenant sheet');
    }
  },

  /**
   * Add sample data to Guest Rooms sheet
   * @private
   */
  _addGuestRoomsSampleData(spreadsheet) {
    const guestRoomsSheet = spreadsheet.getSheetByName(SHEET_NAMES.GUEST_ROOMS);
    if (guestRoomsSheet && guestRoomsSheet.getLastRow() === 1) {
      const guestRoomsData = [
        ['BK001', '201', 'Garden View Suite', 'Deluxe', '2', 'WiFi, TV, Mini-fridge, Private Bath', '$85', '$550', '$2000', 'Occupied', '2024-08-01', '', '2024-08-03', '2024-08-06', '3', '2', 'Sarah Wilson', 'Business Trip', 'Late checkout requested', 'Website', '$255', 'Paid', 'Confirmed', 'Guest requested extra towels'],
        ['BK002', '202', 'City View Room', 'Standard', '1', 'WiFi, TV, Shared Bath', '$65', '$400', '$1500', 'Available', '2024-08-02', 'AC needs maintenance', '', '', '', '', '', '', '', 'Phone', '', '', '', 'Ready for next guest'],
        ['BK003', '203', 'Executive Suite', 'Premium', '4', 'WiFi, TV, Kitchenette, Private Bath, Balcony', '$120', '$750', '$2800', 'Reserved', '2024-08-01', '', '2024-08-10', '2024-08-17', '7', '3', 'Robert Chen', 'Family Vacation', 'Crib needed', 'Email', '$840', 'Deposit Paid', 'Confirmed', 'Anniversary celebration - arrange flowers'],
        ['BK004', '204', 'Economy Room', 'Standard', '1', 'WiFi, TV, Shared Bath', '$55', '$350', '$1200', 'Cleaning', '2024-08-03', '', '', '', '', '', '', '', '', 'Walk-in', '', '', 'Tentative', 'Deep cleaning in progress']
      ];
      guestRoomsSheet.getRange(2, 1, guestRoomsData.length, guestRoomsData[0].length).setValues(guestRoomsData);
      console.log('✅ Sample data added to Guest Rooms sheet');
    }
  },

  /**
   * Add sample data to Budget sheet with negative expenses
   * @private
   */
  _addBudgetSampleData(spreadsheet) {
    const budgetSheet = spreadsheet.getSheetByName(SHEET_NAMES.BUDGET);
    if (budgetSheet && budgetSheet.getLastRow() === 1) {
      const budgetData = [
        // Income entries (positive amounts)
        ['2024-08-01', 'Income', 'Rent Payment - Room 101', '$1150', 'Rent', 'Bank Transfer', 'TXN-001', 'John Smith', 'Receipt-001'],
        ['2024-08-01', 'Income', 'Rent Payment - Room 102', '$1250', 'Rent', 'Bank Transfer', 'TXN-004', 'Maria Garcia', 'Receipt-004'],
        ['2024-08-01', 'Income', 'Rent Payment - Room 103', '$1400', 'Rent', 'Bank Transfer', 'TXN-005', 'David Johnson', 'Receipt-005'],
        ['2024-08-03', 'Income', 'Guest Payment - Room 201', '$255', 'Guest Revenue', 'Cash', 'TXN-003', 'Sarah Wilson', 'Receipt-003'],
        ['2024-08-10', 'Income', 'Guest Deposit - Room 203', '$840', 'Guest Revenue', 'Credit Card', 'TXN-006', 'Robert Chen', 'Receipt-006'],
        
        // Expense entries (negative amounts)
        ['2024-08-02', 'Expense', 'Plumbing Repair - Room 202', '-$175', 'Maintenance', 'Credit Card', 'TXN-002', '', 'Receipt-002'],
        ['2024-08-05', 'Expense', 'Electricity Bill - August', '-$320', 'Utilities', 'Bank Transfer', 'TXN-007', '', 'Receipt-007'],
        ['2024-08-07', 'Expense', 'Cleaning Supplies', '-$85', 'Supplies', 'Cash', 'TXN-008', '', 'Receipt-008'],
        ['2024-08-12', 'Expense', 'Property Insurance - Monthly', '-$450', 'Insurance', 'Bank Transfer', 'TXN-009', '', 'Receipt-009'],
        ['2024-08-15', 'Expense', 'HVAC Maintenance - Room 202', '-$225', 'Maintenance', 'Credit Card', 'TXN-010', '', 'Receipt-010'],
        ['2024-08-18', 'Expense', 'Internet Service - Monthly', '-$120', 'Utilities', 'Bank Transfer', 'TXN-011', '', 'Receipt-011'],
        ['2024-08-20', 'Expense', 'Marketing - Online Ads', '-$150', 'Marketing', 'Credit Card', 'TXN-012', '', 'Receipt-012']
      ];
      budgetSheet.getRange(2, 1, budgetData.length, budgetData[0].length).setValues(budgetData);
      console.log('✅ Sample data added to Budget sheet with negative expenses');
    }
  },

  /**
   * Add sample data to Maintenance Requests sheet
   * @private
   */
  _addMaintenanceSampleData(spreadsheet) {
    const maintenanceSheet = spreadsheet.getSheetByName(SHEET_NAMES.MAINTENANCE);
    if (maintenanceSheet && maintenanceSheet.getLastRow() === 1) {
      const maintenanceData = [
        ['MR-001', '2024-08-01 10:30:00', 'Room 102', 'Plumbing', 'High', 'Kitchen sink is leaking under the cabinet', 'Maria Garcia', 'maria.garcia@email.com', 'Mike Johnson', 'Completed', '$150', '$175', '2024-08-01', '2024-08-01', 'Pipe fitting, Plumber putty', '2', 'photo1.jpg', 'Fixed leak and checked all connections'],
        ['MR-002', '2024-08-02 14:15:00', 'Room 201', 'HVAC', 'Medium', 'Air conditioning not cooling properly', 'Sarah Wilson', 'sarah.wilson@email.com', 'Tom Rodriguez', 'Completed', '$200', '$225', '2024-08-02', '2024-08-03', 'Refrigerant, Filter', '3', '', 'Replaced filter and added refrigerant'],
        ['MR-003', '2024-08-03 09:00:00', 'Common Area', 'Electrical', 'Low', 'Hallway light bulb needs replacement', 'Staff', 'maintenance@whitehouse.com', 'Mike Johnson', 'Completed', '$20', '$15', '2024-08-03', '2024-08-03', 'LED bulb', '0.25', '', 'Quick replacement - working perfectly'],
        ['MR-004', '2024-08-05 16:45:00', 'Room 105', 'Structural', 'Medium', 'Wall needs painting after water damage repair', 'Property Manager', 'manager@whitehouse.com', 'Paint Crew', 'In Progress', '$300', '', '2024-08-06', '', 'Paint, Primer, Brushes', '8', '', 'Prep work completed, painting in progress'],
        ['MR-005', '2024-08-07 11:20:00', 'Room 204', 'Cleaning', 'High', 'Deep cleaning required before next guest', 'Housekeeping', 'housekeeping@whitehouse.com', 'Cleaning Team', 'Pending', '$100', '', '', '', 'Cleaning supplies', '4', '', 'Scheduled for tomorrow morning']
      ];
      maintenanceSheet.getRange(2, 1, maintenanceData.length, maintenanceData[0].length).setValues(maintenanceData);
      console.log('✅ Sample data added to Maintenance Requests sheet');
    }
  }
};
