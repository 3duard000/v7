/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Dashboard Manager - Dashboard.gs
 * 
 * Simplified and user-friendly dashboards with clean layouts and focused metrics.
 */

const Dashboard = {

  /**
   * Create simplified management dashboard
   */
  createManagementDashboard() {
    try {
      console.log('Creating Management Dashboard...');
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let dashboardSheet = spreadsheet.getSheetByName('📊 Management Dashboard');
      
      if (dashboardSheet) {
        dashboardSheet.clear();
      } else {
        dashboardSheet = spreadsheet.insertSheet('📊 Management Dashboard');
      }
      
      dashboardSheet.activate();
      spreadsheet.moveActiveSheet(2);
      
      this._buildSimplifiedManagementDashboard(dashboardSheet);
      
      console.log('✅ Management Dashboard created successfully');
      return 'Management Dashboard created successfully!';
      
    } catch (error) {
      console.error('Error creating Management Dashboard:', error);
      throw new Error('Failed to create Management Dashboard: ' + error.message);
    }
  },

  /**
   * Create simplified financial dashboard
   */
  createFinancialDashboard() {
    try {
      console.log('Creating Financial Dashboard...');
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let financialSheet = spreadsheet.getSheetByName('💰 Financial Dashboard');
      
      if (financialSheet) {
        financialSheet.clear();
      } else {
        financialSheet = spreadsheet.insertSheet('💰 Financial Dashboard');
      }
      
      financialSheet.activate();
      spreadsheet.moveActiveSheet(3);
      
      this._buildSimplifiedFinancialDashboard(financialSheet);
      
      console.log('✅ Financial Dashboard created successfully');
      return 'Financial Dashboard created successfully!';
      
    } catch (error) {
      console.error('Error creating Financial Dashboard:', error);
      throw new Error('Failed to create Financial Dashboard: ' + error.message);
    }
  },

  /**
   * Refresh all dashboards
   */
  refreshAllDashboards() {
    try {
      console.log('Refreshing all dashboards...');
      this.createManagementDashboard();
      this.createFinancialDashboard();
      return 'All dashboards refreshed successfully!';
    } catch (error) {
      console.error('Error refreshing dashboards:', error);
      throw new Error('Failed to refresh dashboards: ' + error.message);
    }
  },

  /**
   * Build simplified management dashboard with clean layout
   * @private
   */
  _buildSimplifiedManagementDashboard(sheet) {
    // Set column widths for better readability
    sheet.setColumnWidths(1, 8, 140);

    let row = 1;

    // Title Section
    sheet.getRange(row, 1, 1, 8).merge();
    sheet.getRange(row, 1)
      .setValue('🏠 WHITE HOUSE MANAGEMENT OVERVIEW')
      .setFontSize(18)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#1c4587')
      .setFontColor('white');
    row += 2;

    // Last Updated
    sheet.getRange(row, 1, 1, 4).merge();
    sheet.getRange(row, 1)
      .setValue(`📅 Last Updated: ${new Date().toLocaleString()}`)
      .setFontStyle('italic')
      .setFontColor('#666666');
    row += 3;

    // Key Metrics Cards
    row = this._addMetricsCards(sheet, row);
    row += 2;

    // Current Occupancy Status
    row = this._addOccupancyStatus(sheet, row);
    row += 2;

    // Payment Status Overview
    row = this._addPaymentOverview(sheet, row);

    sheet.setFrozenRows(5);
  },

  /**
   * Build simplified financial dashboard with clean layout
   * @private
   */
  _buildSimplifiedFinancialDashboard(sheet) {
    // Set column widths
    sheet.setColumnWidths(1, 6, 160);

    let row = 1;

    // Title Section
    sheet.getRange(row, 1, 1, 6).merge();
    sheet.getRange(row, 1)
      .setValue('💰 WHITE HOUSE FINANCIAL OVERVIEW')
      .setFontSize(18)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#0d5016')
      .setFontColor('white');
    row += 2;

    // Last Updated
    sheet.getRange(row, 1, 1, 3).merge();
    sheet.getRange(row, 1)
      .setValue(`📅 Last Updated: ${new Date().toLocaleString()}`)
      .setFontStyle('italic')
      .setFontColor('#666666');
    row += 3;

    // Financial Summary Cards
    row = this._addFinancialCards(sheet, row);
    row += 2;

    // Income vs Expenses Breakdown
    row = this._addIncomeExpenseBreakdown(sheet, row);
    row += 2;

    // Maintenance Status
    row = this._addMaintenanceOverview(sheet, row);

    sheet.setFrozenRows(5);
  },

  /**
   * Add key metrics cards
   * @private
   */
  _addMetricsCards(sheet, startRow) {
    const tenantData = this._getTenantData();
    const guestData = this._getGuestRoomData();
    
    console.log('Tenant data count:', tenantData.length);
    console.log('Guest data count:', guestData.length);
    
    // Calculate key metrics
    const totalRooms = tenantData.length + guestData.length;
    const occupiedRooms = tenantData.filter(t => t.status === 'Occupied').length + 
                         guestData.filter(g => g.status === 'Occupied').length;
    const occupancyRate = totalRooms > 0 ? Math.round((occupiedRooms / totalRooms) * 100) : 0;
    
    const currentTenants = tenantData.filter(t => t.paymentStatus === 'Current').length;
    const lateTenants = tenantData.filter(t => t.paymentStatus === 'Late' || t.paymentStatus === 'Overdue').length;
    
    const totalRent = tenantData.reduce((sum, t) => {
      const rent = this._parseAmount(t.negotiatedPrice || t.rentalPrice);
      console.log(`Room ${t.roomNumber}: rent = ${rent}`);
      return sum + rent;
    }, 0);
    
    const guestRevenue = guestData.reduce((sum, g) => {
      const revenue = this._parseAmount(g.totalAmount);
      console.log(`Guest room ${g.roomNumber}: revenue = ${revenue}`);
      return sum + revenue;
    }, 0);

    console.log('Calculated metrics:', {
      totalRooms,
      occupiedRooms, 
      occupancyRate,
      currentTenants,
      lateTenants,
      totalRent,
      guestRevenue
    });

    // Card 1: Occupancy
    this._createMetricCard(sheet, startRow, 1, 2, 'OCCUPANCY RATE', `${occupancyRate}%`, 
      occupancyRate >= 80 ? '#d9ead3' : occupancyRate >= 60 ? '#fff2cc' : '#f4cccc',
      `${occupiedRooms}/${totalRooms} rooms occupied`);

    // Card 2: Tenant Payments
    this._createMetricCard(sheet, startRow, 3, 4, 'PAYMENT STATUS', `${currentTenants} Current`, 
      lateTenants === 0 ? '#d9ead3' : lateTenants <= 1 ? '#fff2cc' : '#f4cccc',
      `${lateTenants} Late/Overdue`);

    // Card 3: Monthly Rent
    this._createMetricCard(sheet, startRow, 5, 6, 'MONTHLY RENT', `${totalRent.toLocaleString()}`, 
      '#cfe2f3', 'From tenant contracts');

    // Card 4: Guest Revenue
    this._createMetricCard(sheet, startRow, 7, 8, 'GUEST REVENUE', `${guestRevenue.toLocaleString()}`, 
      '#d0e0e3', 'Current bookings');

    return startRow + 4;
  },

  /**
   * Create a metric card with better value display
   * @private
   */
  _createMetricCard(sheet, row, startCol, endCol, title, value, color, subtitle) {
    const cardWidth = endCol - startCol + 1;
    
    // Title row
    const titleRange = sheet.getRange(row, startCol, 1, cardWidth);
    titleRange.merge();
    titleRange.setValue(title)
      .setFontSize(10)
      .setFontWeight('bold')
      .setFontColor('#666666')
      .setHorizontalAlignment('center')
      .setBackground('#f8f9fa')
      .setBorder(true, true, false, true, false, false, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    
    // Value row - this is the main display
    const valueRange = sheet.getRange(row + 1, startCol, 1, cardWidth);
    valueRange.merge();
    valueRange.setValue(value)
      .setFontSize(18)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(color)
      .setBorder(false, true, false, true, false, false, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    
    // Subtitle row
    const subtitleRange = sheet.getRange(row + 2, startCol, 1, cardWidth);
    subtitleRange.merge();
    subtitleRange.setValue(subtitle)
      .setFontSize(9)
      .setFontColor('#666666')
      .setHorizontalAlignment('center')
      .setBackground('#f8f9fa')
      .setBorder(false, true, true, true, false, false, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  },

  /**
   * Add occupancy status table
   * @private
   */
  _addOccupancyStatus(sheet, startRow) {
    // Section header
    sheet.getRange(startRow, 1, 1, 8).merge();
    sheet.getRange(startRow, 1)
      .setValue('🏠 CURRENT ROOM STATUS')
      .setFontSize(14)
      .setFontWeight('bold')
      .setBackground('#e1d5e7')
      .setHorizontalAlignment('center');
    startRow++;

    // Table headers
    const headers = ['Room', 'Type', 'Occupant', 'Status', 'Rate', 'Payment', 'Check-Out', 'Notes'];
    sheet.getRange(startRow, 1, 1, 8).setValues([headers]);
    sheet.getRange(startRow, 1, 1, 8)
      .setFontWeight('bold')
      .setBackground('#f3f3f3')
      .setHorizontalAlignment('center');
    startRow++;

    // Combine tenant and guest data
    const tenantData = this._getTenantData();
    const guestData = this._getGuestRoomData();
    
    const allRooms = [];
    
    // Add tenant rooms
    tenantData.forEach(tenant => {
      allRooms.push([
        tenant.roomNumber,
        'Tenant',
        tenant.name || 'Vacant',
        tenant.status,
        tenant.negotiatedPrice || tenant.rentalPrice || '',
        tenant.paymentStatus || '',
        tenant.status === 'Occupied' ? 'Long-term' : '',
        ''
      ]);
    });
    
    // Add guest rooms
    guestData.forEach(guest => {
      allRooms.push([
        guest.roomNumber,
        'Guest',
        guest.currentGuest || 'Available',
        guest.status,
        guest.dailyRate || '',
        guest.paymentStatus || '',
        guest.checkOutDate || '',
        guest.status === 'Maintenance' ? 'Needs service' : ''
      ]);
    });

    // Sort by room number
    allRooms.sort((a, b) => a[0].toString().localeCompare(b[0].toString()));

    if (allRooms.length > 0) {
      sheet.getRange(startRow, 1, allRooms.length, 8).setValues(allRooms);
      
      // Apply conditional formatting
      for (let i = 0; i < allRooms.length; i++) {
        const currentRow = startRow + i;
        const status = allRooms[i][3];
        const statusCell = sheet.getRange(currentRow, 4);
        
        switch (status) {
          case 'Occupied':
            statusCell.setBackground('#d9ead3');
            break;
          case 'Available':
            statusCell.setBackground('#cfe2f3');
            break;
          case 'Maintenance':
            statusCell.setBackground('#f4cccc');
            break;
          case 'Vacant':
            statusCell.setBackground('#f3f3f3');
            break;
        }
      }
    }

    return startRow + allRooms.length;
  },

  /**
   * Add payment overview
   * @private
   */
  _addPaymentOverview(sheet, startRow) {
    const tenantData = this._getTenantData();
    
    // Section header
    sheet.getRange(startRow, 1, 1, 6).merge();
    sheet.getRange(startRow, 1)
      .setValue('💳 PAYMENT STATUS SUMMARY')
      .setFontSize(14)
      .setFontWeight('bold')
      .setBackground('#fff2cc')
      .setHorizontalAlignment('center');
    startRow++;

    // Payment summary
    const currentCount = tenantData.filter(t => t.paymentStatus === 'Current').length;
    const lateCount = tenantData.filter(t => t.paymentStatus === 'Late').length;
    const overdueCount = tenantData.filter(t => t.paymentStatus === 'Overdue').length;
    const totalTenants = tenantData.filter(t => t.status === 'Occupied').length;

    const summaryData = [
      ['Status', 'Count', 'Percentage', '', 'Action Required', ''],
      ['✅ Current', currentCount, totalTenants > 0 ? `${Math.round((currentCount/totalTenants)*100)}%` : '0%', '', 'None', ''],
      ['⚠️ Late', lateCount, totalTenants > 0 ? `${Math.round((lateCount/totalTenants)*100)}%` : '0%', '', lateCount > 0 ? 'Send reminders' : 'None', ''],
      ['❌ Overdue', overdueCount, totalTenants > 0 ? `${Math.round((overdueCount/totalTenants)*100)}%` : '0%', '', overdueCount > 0 ? 'Immediate action' : 'None', '']
    ];

    sheet.getRange(startRow, 1, summaryData.length, 6).setValues(summaryData);
    
    // Format header row
    sheet.getRange(startRow, 1, 1, 6)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Color code the status column
    for (let i = 1; i < summaryData.length; i++) {
      const row = startRow + i;
      if (summaryData[i][0].includes('Current')) {
        sheet.getRange(row, 1, 1, 6).setBackground('#d9ead3');
      } else if (summaryData[i][0].includes('Late')) {
        sheet.getRange(row, 1, 1, 6).setBackground('#fff2cc');
      } else if (summaryData[i][0].includes('Overdue')) {
        sheet.getRange(row, 1, 1, 6).setBackground('#f4cccc');
      }
    }

    return startRow + summaryData.length;
  },

  /**
   * Add financial summary cards
   * @private
   */
  _addFinancialCards(sheet, startRow) {
    const budgetData = this._getBudgetData();
    
    console.log('Budget data count:', budgetData.length);
    console.log('Budget data sample:', budgetData.slice(0, 3));
    
    const totalIncome = budgetData.filter(b => b.type === 'Income')
      .reduce((sum, item) => {
        const amount = this._parseAmount(item.amount);
        console.log(`Income item: ${item.category} = ${amount}`);
        return sum + amount;
      }, 0);
      
    const totalExpenses = budgetData.filter(b => b.type === 'Expense')
      .reduce((sum, item) => {
        const amount = this._parseAmount(item.amount);
        console.log(`Expense item: ${item.category} = ${amount}`);
        return sum + amount;
      }, 0);
      
    const netIncome = totalIncome + totalExpenses;

    console.log('Financial metrics:', { totalIncome, totalExpenses, netIncome });

    // Income Card
    this._createMetricCard(sheet, startRow, 1, 2, 'TOTAL INCOME', `${totalIncome.toLocaleString()}`, 
      '#d9ead3', 'All revenue sources');

    // Expenses Card
    this._createMetricCard(sheet, startRow, 3, 4, 'TOTAL EXPENSES', `${totalExpenses.toLocaleString()}`, 
      '#f4cccc', 'All operational costs');

    // Net Income Card
    this._createMetricCard(sheet, startRow, 5, 6, 'NET INCOME', `${netIncome.toLocaleString()}`, 
      netIncome >= 0 ? '#cfe2f3' : '#f4cccc', 
      netIncome >= 0 ? 'Profitable' : 'Loss');

    return startRow + 4;
  },

  /**
   * Add income vs expenses breakdown
   * @private
   */
  _addIncomeExpenseBreakdown(sheet, startRow) {
    const budgetData = this._getBudgetData();
    
    // Section header
    sheet.getRange(startRow, 1, 1, 6).merge();
    sheet.getRange(startRow, 1)
      .setValue('📊 INCOME & EXPENSE BREAKDOWN')
      .setFontSize(14)
      .setFontWeight('bold')
      .setBackground('#e1d5e7')
      .setHorizontalAlignment('center');
    startRow++;

    // Group by category
    const categories = {};
    budgetData.forEach(item => {
      if (!categories[item.category]) {
        categories[item.category] = { income: 0, expenses: 0 };
      }
      const amount = this._parseAmount(item.amount);
      if (item.type === 'Income') {
        categories[item.category].income += amount;
      } else {
        // For expenses, use absolute value
        categories[item.category].expenses += Math.abs(amount);
      }
    });

    const headers = ['Category', 'Income', 'Expenses', 'Net', 'Status', ''];
    sheet.getRange(startRow, 1, 1, 6).setValues([headers]);
    sheet.getRange(startRow, 1, 1, 6)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    startRow++;

    Object.keys(categories).forEach(category => {
      const data = categories[category];
      const net = data.income - data.expenses; // Now this will be correct
      const status = net > 0 ? '✅ Positive' : net < 0 ? '❌ Negative' : '➖ Break Even';
      
      const rowData = [
        category,
        `${data.income.toLocaleString()}`,
        `${data.expenses.toLocaleString()}`,
        `${net.toLocaleString()}`,
        status,
        ''
      ];
      
      sheet.getRange(startRow, 1, 1, 6).setValues([rowData]);
      
      // Color code based on net result
      const rowRange = sheet.getRange(startRow, 1, 1, 6);
      if (net > 0) {
        rowRange.setBackground('#d9ead3');
      } else if (net < 0) {
        rowRange.setBackground('#f4cccc');
      }
      
      startRow++;
    });

    return startRow;
  },

  /**
   * Add maintenance overview
   * @private
   */
  _addMaintenanceOverview(sheet, startRow) {
    const maintenanceData = this._getMaintenanceData();
    
    // Section header
    sheet.getRange(startRow, 1, 1, 6).merge();
    sheet.getRange(startRow, 1)
      .setValue('🔧 MAINTENANCE STATUS')
      .setFontSize(14)
      .setFontWeight('bold')
      .setBackground('#fce5cd')
      .setHorizontalAlignment('center');
    startRow++;

    const totalRequests = maintenanceData.length;
    const pendingRequests = maintenanceData.filter(m => m.status === 'Pending').length;
    const inProgressRequests = maintenanceData.filter(m => m.status === 'In Progress').length;
    const completedRequests = maintenanceData.filter(m => m.status === 'Completed').length;
    const emergencyRequests = maintenanceData.filter(m => m.priority === 'Emergency').length;

    const maintenanceStats = [
      ['Status', 'Count', 'Priority', '', '', ''],
      ['📋 Total Requests', totalRequests, '', '', '', ''],
      ['⏳ Pending', pendingRequests, pendingRequests > 2 ? '❌ High' : '✅ Normal', '', '', ''],
      ['🔄 In Progress', inProgressRequests, '', '', '', ''],
      ['✅ Completed', completedRequests, '', '', '', ''],
      ['🚨 Emergency', emergencyRequests, emergencyRequests > 0 ? '❌ Attention Needed' : '✅ None', '', '', '']
    ];

    sheet.getRange(startRow, 1, maintenanceStats.length, 6).setValues(maintenanceStats);
    
    // Format and color code
    sheet.getRange(startRow, 1, 1, 6)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    for (let i = 1; i < maintenanceStats.length; i++) {
      const row = startRow + i;
      if (maintenanceStats[i][0].includes('Emergency') && emergencyRequests > 0) {
        sheet.getRange(row, 1, 1, 6).setBackground('#f4cccc');
      } else if (maintenanceStats[i][0].includes('Pending') && pendingRequests > 2) {
        sheet.getRange(row, 1, 1, 6).setBackground('#fff2cc');
      }
    }

    return startRow + maintenanceStats.length;
  },

  /**
   * Get tenant data from sheet
   * @private
   */
  _getTenantData() {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TENANT);
      if (!sheet || sheet.getLastRow() <= 1) return [];
      
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const tenants = [];
      
      const roomNumberCol = headers.indexOf('Room Number');
      const rentalPriceCol = headers.indexOf('Rental Price');
      const negotiatedPriceCol = headers.indexOf('Negotiated Price');
      const nameCol = headers.indexOf('Current Tenant Name');
      const statusCol = headers.indexOf('Room Status');
      const lastPaymentCol = headers.indexOf('Last Payment Date');
      const paymentStatusCol = headers.indexOf('Payment Status');
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[roomNumberCol]) {
          tenants.push({
            roomNumber: row[roomNumberCol],
            rentalPrice: row[rentalPriceCol],
            negotiatedPrice: row[negotiatedPriceCol],
            name: row[nameCol] || 'Vacant',
            status: row[statusCol] || 'Vacant',
            lastPaymentDate: row[lastPaymentCol],
            paymentStatus: row[paymentStatusCol] || 'N/A'
          });
        }
      }
      
      return tenants;
    } catch (error) {
      console.error('Error getting tenant data:', error);
      return [];
    }
  },

  /**
   * Get guest room data from sheet
   * @private
   */
  _getGuestRoomData() {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.GUEST_ROOMS);
      if (!sheet || sheet.getLastRow() <= 1) return [];
      
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const guestRooms = [];
      
      const roomNumberCol = headers.indexOf('Room Number');
      const dailyRateCol = headers.indexOf('Daily Rate');
      const statusCol = headers.indexOf('Status');
      const currentGuestCol = headers.indexOf('Current Guest');
      const checkOutCol = headers.indexOf('Check-Out Date');
      const totalAmountCol = headers.indexOf('Total Amount');
      const paymentStatusCol = headers.indexOf('Payment Status');
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[roomNumberCol]) {
          guestRooms.push({
            roomNumber: row[roomNumberCol],
            dailyRate: row[dailyRateCol],
            status: row[statusCol] || 'Available',
            currentGuest: row[currentGuestCol],
            checkOutDate: row[checkOutCol] ? new Date(row[checkOutCol]).toLocaleDateString() : null,
            totalAmount: row[totalAmountCol],
            paymentStatus: row[paymentStatusCol]
          });
        }
      }
      
      return guestRooms;
    } catch (error) {
      console.error('Error getting guest room data:', error);
      return [];
    }
  },

  /**
   * Get budget data from sheet
   * @private
   */
  _getBudgetData() {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.BUDGET);
      if (!sheet || sheet.getLastRow() <= 1) return [];
      
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const budget = [];
      
      const typeCol = headers.indexOf('Type');
      const amountCol = headers.indexOf('Amount');
      const categoryCol = headers.indexOf('Category');
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[typeCol] && row[amountCol]) {
          budget.push({
            type: row[typeCol],
            amount: row[amountCol],
            category: row[categoryCol] || 'Other'
          });
        }
      }
      
      return budget;
    } catch (error) {
      console.error('Error getting budget data:', error);
      return [];
    }
  },

  /**
   * Get maintenance data from sheet
   * @private
   */
  _getMaintenanceData() {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MAINTENANCE);
      if (!sheet || sheet.getLastRow() <= 1) return [];
      
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const maintenance = [];
      
      const requestIdCol = headers.indexOf('Request ID');
      const priorityCol = headers.indexOf('Priority');
      const statusCol = headers.indexOf('Status');
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[requestIdCol]) {
          maintenance.push({
            requestId: row[requestIdCol],
            priority: row[priorityCol] || 'Medium',
            status: row[statusCol] || 'Pending'
          });
        }
      }
      
      return maintenance;
    } catch (error) {
      console.error('Error getting maintenance data:', error);
      return [];
    }
  },

  /**
   * Parse amount from string
   * @private
   */
  _parseAmount(amount) {
    if (!amount) return 0;
    if (typeof amount === 'number') return amount;
    
    const cleanAmount = amount.toString().replace(/[$,\s]/g, '');
    return parseFloat(cleanAmount) || 0;
  }
};

/**
 * Wrapper functions for menu integration
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
