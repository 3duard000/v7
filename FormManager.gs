/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Form Manager - FormManager.gs
 * 
 * This module handles all Google Forms creation and linking to spreadsheet.
 * Updated with delayed sheet renaming approach.
 */

const FormManager = {

  /**
   * Create all Google Forms and link them to the spreadsheet
   */
  createGoogleForms: function() {
    console.log('Creating Google Forms...');
    
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      // Create Tenant Application Form
      console.log('Creating Tenant Application Form...');
      const applicationForm = this._createTenantApplicationForm();
      applicationForm.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
      
      // Set up a trigger to add "Processed" column when first response is submitted
      this._setupProcessedColumnTrigger(applicationForm);
      
      console.log('✅ Tenant Application Form created and linked: ' + applicationForm.getPublishedUrl());
      
      // Create Move-Out Request Form
      console.log('Creating Move-Out Request Form...');
      const moveOutForm = this._createMoveOutRequestForm();
      moveOutForm.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
      
      console.log('✅ Move-Out Request Form created and linked: ' + moveOutForm.getPublishedUrl());
      
      // Create Guest Check-In Form
      console.log('Creating Guest Check-In Form...');
      const guestCheckInForm = this._createGuestCheckInForm();
      guestCheckInForm.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
      
      console.log('✅ Guest Check-In Form created and linked: ' + guestCheckInForm.getPublishedUrl());
      
      // Store form URLs in the spreadsheet for easy access
      this._storeFormUrls(
        applicationForm.getPublishedUrl(), 
        moveOutForm.getPublishedUrl(), 
        guestCheckInForm.getPublishedUrl()
      );
      
      // Store form IDs for delayed processing
      this._storeFormIds(
        applicationForm.getId(),
        moveOutForm.getId(),
        guestCheckInForm.getId()
      );
      
      console.log('⚠️  IMPORTANT NOTES:');
      console.log('   1. Response sheets will be automatically renamed after form submissions');
      console.log('   2. File upload questions must be added manually to the Tenant Application form');
      console.log('   3. A delayed trigger will handle sheet renaming');
      
    } catch (error) {
      console.error('Error creating forms:', error);
      throw new Error('Failed to create Google Forms: ' + error.message);
    }
  },

  /**
   * Create delayed trigger to rename form response sheets
   * This gets called from the main setup after a delay
   */
  setupDelayedSheetRenaming: function() {
    console.log('Setting up delayed sheet renaming trigger...');
    
    // Create a trigger that runs in 3 minutes to rename sheets
    ScriptApp.newTrigger('performDelayedSheetRenaming')
      .timeBased()
      .after(3 * 60 * 1000) // 3 minutes delay
      .create();
    
    console.log('✅ Delayed renaming trigger set for 3 minutes from now');
  },

  /**
   * Perform the actual sheet renaming after delay
   * This function will be called by the trigger
   */
  performDelayedSheetRenaming: function() {
    try {
      console.log('Starting delayed sheet renaming process...');
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      // First, try to trigger form responses by submitting dummy data
      this._triggerFormResponseSheetCreation();
      
      // Wait a bit more for sheets to be created
      Utilities.sleep(30000); // 30 seconds
      
      // Now attempt to rename the sheets
      this._renameFormResponseSheets();
      
      // Clean up the trigger
      this._cleanupDelayedTriggers();
      
      console.log('✅ Delayed sheet renaming completed');
      
    } catch (error) {
      console.error('Error in delayed sheet renaming:', error);
      this._cleanupDelayedTriggers();
    }
  },

  /**
   * Trigger form response sheet creation by submitting dummy responses
   * @private
   */
  _triggerFormResponseSheetCreation: function() {
    try {
      console.log('Triggering form response sheet creation...');
      
      const formIds = this._getStoredFormIds();
      
      if (!formIds) {
        console.log('No stored form IDs found, attempting manual approach');
        return;
      }
      
      // Submit dummy responses to each form to trigger sheet creation
      formIds.forEach((formId, index) => {
        try {
          const form = FormApp.openById(formId);
          console.log(`Submitting dummy response to: ${form.getTitle()}`);
          
          const formResponse = form.createResponse();
          const items = form.getItems();
          
          // Add dummy responses to required fields only
          items.forEach(item => {
            try {
              const itemType = item.getType();
              
              if (item.isRequired()) {
                if (itemType === FormApp.ItemType.TEXT) {
                  formResponse.withItemResponse(item.asTextItem().createResponse('DUMMY'));
                } else if (itemType === FormApp.ItemType.PARAGRAPH_TEXT) {
                  formResponse.withItemResponse(item.asParagraphTextItem().createResponse('DUMMY DATA'));
                } else if (itemType === FormApp.ItemType.DATE) {
                  formResponse.withItemResponse(item.asDateItem().createResponse(new Date()));
                } else if (itemType === FormApp.ItemType.MULTIPLE_CHOICE) {
                  const choices = item.asMultipleChoiceItem().getChoices();
                  if (choices.length > 0) {
                    formResponse.withItemResponse(item.asMultipleChoiceItem().createResponse(choices[0].getValue()));
                  }
                }
              }
            } catch (itemError) {
              console.log(`Skipped item in form ${index}: ${itemError.message}`);
            }
          });
          
          // Submit the response
          formResponse.submit();
          console.log(`✅ Dummy response submitted to form ${index + 1}`);
          
          // Small delay between submissions
          Utilities.sleep(2000);
          
        } catch (formError) {
          console.error(`Error submitting to form ${index}: ${formError.message}`);
        }
      });
      
    } catch (error) {
      console.error('Error triggering form responses:', error);
    }
  },

  /**
   * Rename form response sheets with proper identification
   * @private
   */
  _renameFormResponseSheets: function() {
    try {
      console.log('Starting form response sheet renaming...');
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheets = spreadsheet.getSheets();
      
      sheets.forEach(sheet => {
        const sheetName = sheet.getName();
        
        // Only process Form Responses sheets
        if (!sheetName.startsWith('Form Responses')) {
          return;
        }
        
        // Skip if no data (no headers)
        if (sheet.getLastColumn() === 0) {
          console.log(`Skipping empty sheet: ${sheetName}`);
          return;
        }
        
        try {
          console.log(`Analyzing sheet: ${sheetName}`);
          
          // Get headers to identify the form type
          const headers = sheet.getRange(1, 1, 1, Math.min(15, sheet.getLastColumn())).getValues()[0];
          const headerString = headers.join(' ').toLowerCase();
          
          console.log(`Headers preview: ${headers.slice(0, 5).join(', ')}...`);
          
          let newName = null;
          
          // Identify sheet type by analyzing headers
          if (this._isTenantApplicationSheet(headerString, headers)) {
            newName = 'Tenant Application';
          } else if (this._isMoveOutRequestSheet(headerString, headers)) {
            newName = 'Move-Out Requests';
          } else if (this._isGuestCheckInSheet(headerString, headers)) {
            newName = 'Guest Check-Ins';
          }
          
          if (newName) {
            // Check for name conflicts and resolve them
            let finalName = newName;
            let counter = 1;
            while (spreadsheet.getSheetByName(finalName)) {
              if (spreadsheet.getSheetByName(finalName) === sheet) {
                // Same sheet, no need to rename
                console.log(`Sheet "${sheetName}" already has correct name "${finalName}"`);
                return;
              }
              finalName = `${newName} ${counter}`;
              counter++;
            }
            
            // Rename the sheet
            sheet.setName(finalName);
            console.log(`✅ Renamed "${sheetName}" to "${finalName}"`);
            
            // Remove dummy responses if they exist
            this._removeDummyResponses(sheet);
            
          } else {
            console.log(`Could not identify form type for: ${sheetName}`);
          }
          
        } catch (analysisError) {
          console.error(`Error analyzing sheet ${sheetName}: ${analysisError.message}`);
        }
      });
      
    } catch (error) {
      console.error('Error in renaming process:', error);
    }
  },

  /**
   * Check if sheet contains tenant application headers
   * @private
   */
  _isTenantApplicationSheet: function(headerString, headers) {
    const tenantKeywords = [
      'full name',
      'monthly income',
      'employment status',
      'current address',
      'employer',
      'proof of income'
    ];
    
    let matchCount = 0;
    tenantKeywords.forEach(keyword => {
      if (headerString.includes(keyword)) {
        matchCount++;
      }
    });
    
    // Need at least 3 matches to be confident
    return matchCount >= 3;
  },

  /**
   * Check if sheet contains move-out request headers
   * @private
   */
  _isMoveOutRequestSheet: function(headerString, headers) {
    const moveOutKeywords = [
      'tenant name',
      'planned move-out date',
      'move-out date',
      'forwarding address',
      'satisfaction',
      'reason for moving'
    ];
    
    let matchCount = 0;
    moveOutKeywords.forEach(keyword => {
      if (headerString.includes(keyword)) {
        matchCount++;
      }
    });
    
    return matchCount >= 2;
  },

  /**
   * Check if sheet contains guest check-in headers
   * @private
   */
  _isGuestCheckInSheet: function(headerString, headers) {
    const guestKeywords = [
      'guest name',
      'check-in date',
      'number of nights',
      'number of guests',
      'purpose of visit'
    ];
    
    let matchCount = 0;
    guestKeywords.forEach(keyword => {
      if (headerString.includes(keyword)) {
        matchCount++;
      }
    });
    
    return matchCount >= 3;
  },

  /**
   * Remove dummy responses from sheets
   * @private
   */
  _removeDummyResponses: function(sheet) {
    try {
      if (sheet.getLastRow() <= 1) return;
      
      const data = sheet.getDataRange().getValues();
      const rowsToDelete = [];
      
      // Find rows with dummy data
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowString = row.join(' ').toLowerCase();
        
        if (rowString.includes('dummy') || rowString.includes('test')) {
          rowsToDelete.push(i + 1); // Store 1-based row numbers
        }
      }
      
      // Delete dummy rows (start from the end to maintain row numbers)
      rowsToDelete.reverse().forEach(rowNum => {
        sheet.deleteRow(rowNum);
        console.log(`Removed dummy response from row ${rowNum}`);
      });
      
    } catch (error) {
      console.error('Error removing dummy responses:', error);
    }
  },

  /**
   * Store form IDs for later use
   * @private
   */
  _storeFormIds: function(tenantFormId, moveOutFormId, guestFormId) {
    try {
      const properties = PropertiesService.getScriptProperties();
      properties.setProperties({
        'TENANT_FORM_ID': tenantFormId,
        'MOVEOUT_FORM_ID': moveOutFormId,
        'GUEST_FORM_ID': guestFormId
      });
      console.log('Form IDs stored in script properties');
    } catch (error) {
      console.error('Error storing form IDs:', error);
    }
  },

  /**
   * Get stored form IDs
   * @private
   */
  _getStoredFormIds: function() {
    try {
      const properties = PropertiesService.getScriptProperties();
      const formIds = [
        properties.getProperty('TENANT_FORM_ID'),
        properties.getProperty('MOVEOUT_FORM_ID'),
        properties.getProperty('GUEST_FORM_ID')
      ];
      
      // Return only if all IDs exist
      if (formIds.every(id => id)) {
        return formIds;
      }
      
      return null;
    } catch (error) {
      console.error('Error getting stored form IDs:', error);
      return null;
    }
  },

  /**
   * Clean up delayed triggers
   * @private
   */
  _cleanupDelayedTriggers: function() {
    try {
      const triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'performDelayedSheetRenaming') {
          ScriptApp.deleteTrigger(trigger);
          console.log('Cleaned up delayed renaming trigger');
        }
      });
    } catch (error) {
      console.error('Error cleaning up triggers:', error);
    }
  },

  /**
   * Set up a trigger to add "Processed" column when first response is submitted
   * @private
   */
  _setupProcessedColumnTrigger: function(form) {
    try {
      // Create a form submit trigger that will add the Processed column
      ScriptApp.newTrigger('onTenantApplicationSubmit')
        .forForm(form)
        .onFormSubmit()
        .create();
      
      console.log('✅ Set up trigger to add "Processed" column on form submission');
      
    } catch (error) {
      console.error('Error setting up Processed column trigger:', error);
    }
  },

  // ... [Include all the existing form creation methods: _createTenantApplicationForm, _createMoveOutRequestForm, _createGuestCheckInForm, _storeFormUrls - keep them exactly as they are]

  /**
   * Create the Tenant Application Google Form with clean field types
   * @private
   */
  _createTenantApplicationForm: function() {
    const form = FormApp.create('🏠 ' + EMAIL_CONFIG.PROPERTY_NAME + ' - Tenant Application');
    
    // Basic form settings
    form.setDescription('Thank you for your interest in ' + EMAIL_CONFIG.PROPERTY_NAME + '! Please complete this application form to apply for a room. All required fields must be completed for your application to be processed.');
    
    // PERSONAL INFORMATION SECTION
    form.addSectionHeaderItem()
      .setTitle('Personal Information')
      .setHelpText('Please provide your basic contact details.');
    
    // Full Name - SHORT ANSWER (Required)
    form.addTextItem()
      .setTitle('Full Name')
      .setRequired(true)
      .setHelpText('Enter your first and last name as it appears on your ID');
    
    // Email Address - SHORT ANSWER (Required)
    form.addTextItem()
      .setTitle('Email Address')
      .setRequired(true)
      .setHelpText('We will use this email for all communication regarding your application');
    
    // Phone Number - SHORT ANSWER (Required)
    form.addTextItem()
      .setTitle('Phone Number')
      .setRequired(true)
      .setHelpText('Include area code (e.g., 555-123-4567)');
    
    // Current Address - PARAGRAPH (Required)
    form.addParagraphTextItem()
      .setTitle('Current Address')
      .setRequired(true)
      .setHelpText('Please provide your complete current address including street, city, state, and zip code');
    
    // HOUSING PREFERENCES SECTION
    form.addSectionHeaderItem()
      .setTitle('Housing Preferences')
      .setHelpText('Tell us about your housing needs and preferences.');
    
    // Desired Move-in Date - DATE (Required)
    form.addDateItem()
      .setTitle('Desired Move-in Date')
      .setRequired(true)
      .setHelpText('What date would you like to move in?');
    
    // Preferred Room - SHORT ANSWER (Optional)
    form.addTextItem()
      .setTitle('Preferred Room')
      .setRequired(false)
      .setHelpText('If you have a specific room preference, enter it here (e.g., Room 101). Leave blank for any available room.');
    
    // EMPLOYMENT & FINANCIAL SECTION
    form.addSectionHeaderItem()
      .setTitle('Employment & Financial Information')
      .setHelpText('We need to verify your ability to pay rent. All information is kept confidential.');
    
    // Employment Status - SHORT ANSWER (Required)
    form.addTextItem()
      .setTitle('Employment Status')
      .setRequired(true)
      .setHelpText('e.g., Full-time employed, Part-time employed, Student, Self-employed, Unemployed');
    
    // Employer/School Name - SHORT ANSWER (Required)
    form.addTextItem()
      .setTitle('Employer/School Name')
      .setRequired(true)
      .setHelpText('Name of your current employer or educational institution');
    
    // Monthly Income - SHORT ANSWER (Required)
    form.addTextItem()
      .setTitle('Monthly Income')
      .setRequired(true)
      .setHelpText('Enter your gross monthly income (e.g., $3,500)');
    
    // Proof of Income - TEXT (Will need manual conversion to FILE UPLOAD)
    form.addTextItem()
      .setTitle('Proof of Income')
      .setRequired(true)
      .setHelpText('⚠️ FORM ADMIN: Please manually convert this to a File Upload question and set appropriate file restrictions (PDF, DOC, JPG, PNG). Allow up to 3 files.')
      .setValidation(FormApp.createTextValidation()
        .setHelpText('This field will be converted to file upload - please ignore for now')
        .build());
    
    // REFERENCES & CONTACTS SECTION
    form.addSectionHeaderItem()
      .setTitle('References & Emergency Contact')
      .setHelpText('Please provide at least one reference and an emergency contact.');
    
    // Reference 1 - SHORT ANSWER (Required)
    form.addTextItem()
      .setTitle('Reference 1 (Required)')
      .setRequired(true)
      .setHelpText('Please provide name and phone number in this format: Josh (555-123-4567)');
    
    // Reference 2 - SHORT ANSWER (Optional)
    form.addTextItem()
      .setTitle('Reference 2 (Optional)')
      .setRequired(false)
      .setHelpText('Please provide name and phone number in this format: Sarah (555-987-6543)');
    
    // Emergency Contact - SHORT ANSWER (Optional)
    form.addTextItem()
      .setTitle('Emergency Contact')
      .setRequired(false)
      .setHelpText('Please provide name and phone number in this format: Mary Smith (555-456-7890)');
    
    // ABOUT YOU SECTION
    form.addSectionHeaderItem()
      .setTitle('About You')
      .setHelpText('Help us get to know you better as a potential tenant.');
    
    // Tell us about yourself - PARAGRAPH (Required)
    form.addParagraphTextItem()
      .setTitle('Tell us about yourself')
      .setRequired(true)
      .setHelpText('Please share information about your lifestyle, hobbies, work schedule, or anything else you think would be relevant for us to know as your potential landlord.');
    
    // Special Needs or Requests - SHORT ANSWER (Optional)
    form.addTextItem()
      .setTitle('Special Needs or Requests')
      .setRequired(false)
      .setHelpText('Do you have any accessibility needs, pet requests, or other special requirements?');
    
    // APPLICATION AGREEMENT SECTION
    form.addSectionHeaderItem()
      .setTitle('Application Agreement')
      .setHelpText('Please review and agree to the following terms.');
    
    // Application Agreement - YES/NO CHOICE (Required)
    const agreementItem = form.addMultipleChoiceItem()
      .setTitle('Application Agreement')
      .setRequired(true);
    agreementItem.setChoices([
        agreementItem.createChoice('Yes, I agree'),
        agreementItem.createChoice('No, I do not agree')
      ])
      .setHelpText('By selecting "Yes, I agree", you confirm that: (1) All information provided is true and accurate, (2) You understand that false information may result in application denial, (3) You consent to background and credit checks, (4) You understand that application review may take up to 48-72 hours, and (5) Submission of this application does not guarantee approval or room reservation.');
    
    // Set form confirmation message
    form.setConfirmationMessage(
      '✅ Thank you for submitting your tenant application!\n\n' +
      'Your application has been received and will be reviewed within 48-72 hours. ' +
      'We will contact you via email with our decision.\n\n' +
      'If you have any questions, please don\'t hesitate to contact us.\n\n' +
      'Best regards,\n' + EMAIL_CONFIG.MANAGEMENT_TEAM
    );
    
    return form;
  },

  /**
   * Create the Move-Out Request Google Form
   * @private
   */
  _createMoveOutRequestForm: function() {
    const form = FormApp.create('🏠 ' + EMAIL_CONFIG.PROPERTY_NAME + ' - Move-Out Request');
    
    // Basic form settings
    form.setDescription('We\'re sorry to see you go! Please complete this form to process your move-out request and help us improve our services.');
    
    // TENANT INFORMATION
    form.addSectionHeaderItem()
      .setTitle('Tenant Information')
      .setHelpText('Please confirm your current tenant details.');
    
    form.addTextItem()
      .setTitle('Tenant Name')
      .setRequired(true)
      .setHelpText('Your full name as it appears on the lease');
    
    form.addTextItem()
      .setTitle('Email Address')
      .setRequired(true)
      .setHelpText('Email address on file with us');
    
    form.addTextItem()
      .setTitle('Phone Number')
      .setRequired(true)
      .setHelpText('Your current phone number');
    
    form.addTextItem()
      .setTitle('Room Number')
      .setRequired(true)
      .setHelpText('e.g., Room 101');
    
    // MOVE-OUT DETAILS
    form.addSectionHeaderItem()
      .setTitle('Move-Out Details')
      .setHelpText('Please provide your move-out timeline and forwarding information.');
    
    form.addDateItem()
      .setTitle('Planned Move-Out Date')
      .setRequired(true)
      .setHelpText('What date do you plan to vacate the room?');
    
    form.addParagraphTextItem()
      .setTitle('Forwarding Address')
      .setRequired(true)
      .setHelpText('Complete address where we should send your security deposit and any correspondence');
    
    // FEEDBACK SECTION
    form.addSectionHeaderItem()
      .setTitle('Feedback (Optional)')
      .setHelpText('Your feedback helps us improve our services for future tenants.');
    
    const reasonItem = form.addMultipleChoiceItem()
      .setTitle('Primary Reason for Moving')
      .setRequired(true);
    reasonItem.setChoices([
        reasonItem.createChoice('Job relocation'),
        reasonItem.createChoice('School/education'),
        reasonItem.createChoice('Financial reasons'),
        reasonItem.createChoice('Found better housing'),
        reasonItem.createChoice('Lifestyle change'),
        reasonItem.createChoice('Property issues'),
        reasonItem.createChoice('Other')
      ]);
    reasonItem.showOtherOption(true);
    
    form.addParagraphTextItem()
      .setTitle('Additional Details')
      .setRequired(false)
      .setHelpText('Any additional details about your reason for moving?');
    
    const satisfactionItem = form.addMultipleChoiceItem()
      .setTitle('Overall Satisfaction')
      .setRequired(true);
    satisfactionItem.setChoices([
        satisfactionItem.createChoice('Very Satisfied'),
        satisfactionItem.createChoice('Satisfied'),
        satisfactionItem.createChoice('Neutral'),
        satisfactionItem.createChoice('Dissatisfied'),
        satisfactionItem.createChoice('Very Dissatisfied')
      ])
      .setHelpText('How would you rate your overall experience living here?');
    
    form.addParagraphTextItem()
      .setTitle('What did you like most about living here?')
      .setRequired(false);
    
    form.addParagraphTextItem()
      .setTitle('What could we improve?')
      .setRequired(false);
    
    const recommendItem = form.addMultipleChoiceItem()
      .setTitle('Would you recommend us to others?')
      .setRequired(true);
    recommendItem.setChoices([
        recommendItem.createChoice('Definitely Yes'),
        recommendItem.createChoice('Probably Yes'),
        recommendItem.createChoice('Neutral'),
        recommendItem.createChoice('Probably No'),
        recommendItem.createChoice('Definitely No')
      ]);
    
    form.setConfirmationMessage(
      '✅ Your move-out request has been submitted!\n\n' +
      'We will process your request and contact you within 24 hours to schedule ' +
      'a move-out inspection and discuss security deposit return.\n\n' +
      'Thank you for being a tenant with us!'
    );
    
    return form;
  },

  /**
   * Create the Guest Check-In Google Form
   * @private
   */
  _createGuestCheckInForm: function() {
    const form = FormApp.create('🏠 ' + EMAIL_CONFIG.PROPERTY_NAME + ' - Guest Check-In');
    
    // Basic form settings
    form.setDescription('Welcome to ' + EMAIL_CONFIG.PROPERTY_NAME + '! Please complete this check-in form to confirm your reservation and provide us with the necessary information for your stay.');
    
    // GUEST INFORMATION
    form.addSectionHeaderItem()
      .setTitle('Guest Information')
      .setHelpText('Please provide your contact details.');
    
    form.addTextItem()
      .setTitle('Guest Name')
      .setRequired(true)
      .setHelpText('Full name of the primary guest');
    
    form.addTextItem()
      .setTitle('Email')
      .setRequired(true)
      .setHelpText('Email address for confirmation and communication');
    
    form.addTextItem()
      .setTitle('Phone')
      .setRequired(true)
      .setHelpText('Phone number in case we need to reach you during your stay');
    
    // STAY DETAILS
    form.addSectionHeaderItem()
      .setTitle('Stay Details')
      .setHelpText('Please provide information about your booking.');
    
    form.addTextItem()
      .setTitle('Room Number')
      .setRequired(true)
      .setHelpText('The room number you have reserved (e.g., 201)');
    
    form.addDateItem()
      .setTitle('Check-In Date')
      .setRequired(true)
      .setHelpText('Date you are checking in');
    
    form.addTextItem()
      .setTitle('Number of Nights')
      .setRequired(true)
      .setHelpText('How many nights will you be staying? (e.g., 3)')
      .setValidation(FormApp.createTextValidation()
        .requireNumber()
        .setHelpText('Please enter a valid number')
        .build());
    
    form.addTextItem()
      .setTitle('Number of Guests')
      .setRequired(true)
      .setHelpText('Total number of people in your party (e.g., 2)')
      .setValidation(FormApp.createTextValidation()
        .requireNumber()
        .setHelpText('Please enter a valid number')
        .build());
    
    // VISIT PURPOSE & REQUESTS
    form.addSectionHeaderItem()
      .setTitle('Visit Details')
      .setHelpText('Help us make your stay more comfortable.');
    
    form.addParagraphTextItem()
      .setTitle('Purpose of Visit')
      .setRequired(true)
      .setHelpText('What brings you to the area? (e.g., business trip, vacation, family visit)');
    
    form.addParagraphTextItem()
      .setTitle('Special Requests')
      .setRequired(false)
      .setHelpText('Any special requests or needs for your stay? (e.g., early check-in, late checkout, extra towels)');
    
    form.setConfirmationMessage(
      '✅ Thank you for checking in!\n\n' +
      'Your check-in form has been submitted successfully. Our staff will process ' +
      'your information and prepare your room.\n\n' +
      'If you have any questions or needs during your stay, please don\'t hesitate ' +
      'to contact us.\n\n' +
      'Welcome and enjoy your stay!'
    );
    
    return form;
  },

  /**
   * Store form URLs in the spreadsheet for easy access
   * @private
   */
  _storeFormUrls: function(applicationUrl, moveOutUrl, guestCheckInUrl) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create or get Form URLs sheet
    let urlSheet = spreadsheet.getSheetByName('Form URLs');
    if (!urlSheet) {
      urlSheet = spreadsheet.insertSheet('Form URLs');
      
      // Set headers and URLs
      urlSheet.getRange('A1').setValue('Form Name');
      urlSheet.getRange('B1').setValue('URL');
      urlSheet.getRange('C1').setValue('Notes');
      
      urlSheet.getRange('A2').setValue('🏠 Tenant Application');
      urlSheet.getRange('B2').setValue(applicationUrl);
      urlSheet.getRange('C2').setValue('⚠️ Manually add File Upload for Proof of Income');
      
      urlSheet.getRange('A3').setValue('🏠 Move-Out Request');
      urlSheet.getRange('B3').setValue(moveOutUrl);
      urlSheet.getRange('C3').setValue('Complete form - no manual changes needed');
      
      urlSheet.getRange('A4').setValue('🏠 Guest Check-In');
      urlSheet.getRange('B4').setValue(guestCheckInUrl);
      urlSheet.getRange('C4').setValue('Complete form - no manual changes needed');
      
      // Format headers
      urlSheet.getRange('A1:C1').setBackground(COLORS.HEADER_BLUE).setFontColor('white').setFontWeight('bold');
      urlSheet.setColumnWidth(1, 200);
      urlSheet.setColumnWidth(2, 400);
      urlSheet.setColumnWidth(3, 300);
    }
    
    console.log('Form URLs stored in "Form URLs" sheet with manual setup notes');
  }
};

/**
 * Global function that will be called by the delayed trigger
 */
function performDelayedSheetRenaming() {
  return FormManager.performDelayedSheetRenaming();
}
