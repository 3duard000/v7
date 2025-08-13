/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Trigger Manager - TriggerManager.gs
 * 
 * This module handles all automated triggers for email scheduling and dashboard updates.
 */

const TriggerManager = {

  /**
   * Set up time-based triggers for automated emails and dashboard updates
   */
  setupTriggers() {
    // Delete existing triggers first
    this._deleteExistingTriggers();
    
    console.log('Setting up triggers...');
    
    // Create daily email check trigger (runs every day at 9:00 AM)
    ScriptApp.newTrigger('checkAndRunDailyTasks')
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();
    
    // Create dashboard refresh triggers (3 times per day)
    // Morning refresh at 8:00 AM
    ScriptApp.newTrigger('refreshAllDashboards')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();
    
    // Afternoon refresh at 2:00 PM  
    ScriptApp.newTrigger('refreshAllDashboards')
      .timeBased()
      .everyDays(1)
      .atHour(14)
      .create();
    
    // Evening refresh at 8:00 PM
    ScriptApp.newTrigger('refreshAllDashboards')
      .timeBased()
      .everyDays(1)
      .atHour(20)
      .create();
    
    console.log('Triggers set up successfully:');
    console.log('- Daily email check at 9:00 AM');
    console.log('- Dashboard refresh at 8:00 AM, 2:00 PM, and 8:00 PM');
  },

  /**
   * Get information about current triggers
   */
  getTriggerInfo() {
    const triggers = ScriptApp.getProjectTriggers();
    const relevantTriggers = triggers.filter(trigger => 
      ['sendRentReminders', 'sendLatePaymentAlerts', 'sendMonthlyInvoices', 'checkAndRunDailyTasks', 'refreshAllDashboards'].includes(trigger.getHandlerFunction())
    );
    
    console.log('Current triggers:');
    relevantTriggers.forEach(trigger => {
      const eventType = trigger.getEventType();
      const handlerFunction = trigger.getHandlerFunction();
      
      if (eventType === ScriptApp.EventType.CLOCK) {
        const triggerSource = trigger.getTriggerSource();
        console.log(`- ${handlerFunction}: ${triggerSource} (Time-based)`);
      } else {
        console.log(`- ${handlerFunction}: ${eventType}`);
      }
    });
    
    return `Found ${relevantTriggers.length} relevant triggers. Check execution transcript for details.`;
  },

  /**
   * Delete existing triggers to avoid duplicates
   * @private
   */
  _deleteExistingTriggers() {
    const triggers = ScriptApp.getProjectTriggers();
    const functionNames = ['sendRentReminders', 'sendLatePaymentAlerts', 'sendMonthlyInvoices', 'checkAndRunDailyTasks', 'refreshAllDashboards'];
    
    let deletedCount = 0;
    triggers.forEach(trigger => {
      if (functionNames.includes(trigger.getHandlerFunction())) {
        ScriptApp.deleteTrigger(trigger);
        console.log(`Deleted existing trigger for ${trigger.getHandlerFunction()}`);
        deletedCount++;
      }
    });
    
    if (deletedCount > 0) {
      console.log(`Deleted ${deletedCount} existing triggers`);
    }
  }
};
