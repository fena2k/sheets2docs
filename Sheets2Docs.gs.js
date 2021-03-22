function onOpen() {
    var createMenu = SpreadsheetApp.getUi().createMenu('Sheets2Docs');
    createMenu.addItem('Generate Patent Doc', 'generatePatentDoc');
    createMenu.addToUi();
  }
  
  function generatePatentDoc() {
    var ui = SpreadsheetApp.getUi();
    var fileName = ui.prompt('Please enter file name:');
    var agentName = ui.prompt('Please enter agent name:');
    var saveFolder = DriveApp.getFolderById('XXXXXX'); // ID of a folder where we want to save the generated Doc file
    var templateDoc = DriveApp.getFileById('1GmN6vFQB9vt8EPhp5eXiYpvuTAKqc465CX7pwxoofkQ').makeCopy(fileName.getResponseText(), saveFolder); // ID of the Patent Doc Template file
      
    var createDoc = DocumentApp.openById(templateDoc.getId()).getBody();
  
    var styleHeading2 = createDoc.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING2);
    styleHeading2[DocumentApp.Attribute.FONT_SIZE] = 14;
    styleHeading2[DocumentApp.Attribute.BOLD] = true;
    createDoc.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING2, styleHeading2);
    
    // 1st page
    var coverSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cover Page');
  
    createDoc.replaceText('{{Title}}', coverSheet.getRange(1, 2).getValue());
    createDoc.replaceText('{{Agent Name}}', agentName.getResponseText());
    createDoc.replaceText('{{Inventor}}', coverSheet.getRange(3, 2).getValue());
    createDoc.replaceText('{{Patent No}}', coverSheet.getRange(2, 2).getValue());
  
    // 2nd page - Table of Contents - can't be inserted
    createDoc.appendPageBreak();
  
    // 3rd page - Abstract
    var abstractSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Abstract');
  
    createDoc.appendParagraph(abstractSheet.getRange(4, 2).getValue()).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    createDoc.appendParagraph("");
    
    var abstractRange = abstractSheet.getRange(4, 2, 2, 1).getDataRegion(SpreadsheetApp.Dimension.ROWS);
    var abstractValues = abstractRange.getValues();
    var abstractBackgrounds = abstractRange.getBackgrounds();
    var abstractStyles = abstractRange.getTextStyles();
    var abstractTable = createDoc.appendTable(abstractValues);
    
    abstractTable.setBorderWidth(1);
  
    // Header cell, font size 11, background color #9fc5e8; Content font Roboto;
  
    for(var i = 0; i < abstractTable.getNumRows(); i++){
      for(var j = 0; j < abstractTable.getRow(i).getNumCells(); j++) {
        var abstractObj = {};
          abstractObj[DocumentApp.Attribute.BACKGROUND_COLOR] = abstractBackgrounds[i][j];
          abstractObj[DocumentApp.Attribute.FONT_SIZE] = abstractStyles[i][j].getFontSize();
          abstractObj[DocumentApp.Attribute.FONT_FAMILY] = abstractStyles[i][j].getFontFamily();
          if(abstractStyles[i][j].isBold()) {
            abstractObj[DocumentApp.Attribute.BOLD] = true;
          }
          abstractTable.getRow(i).getCell(j).setAttributes(abstractObj);
      }
    }
    
    createDoc.appendPageBreak();
  
    // 4th page - Relevant Classification codes
    var codesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Classification Codes');
  
    createDoc.appendParagraph(codesSheet.getRange(1, 2).getValue()).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    createDoc.appendParagraph("");
  
    // Header cell background color #9fc5e8
  
    var codesRange = codesSheet.getRange(3, 2, 6, 2).getDataRegion(SpreadsheetApp.Dimension.ROWS);
    var codesValues = codesRange.getValues();
    var codesBackgrounds = codesRange.getBackgrounds();
    var codesStyles = codesRange.getTextStyles();
    var codesTable = createDoc.appendTable(codesValues);
  
    codesTable.setBorderWidth(1);
  
    for(var i = 0; i < codesTable.getNumRows(); i++) {
      for(var j = 0; j < codesTable.getRow(i).getNumCells(); j++) {
        var codesObj = {};
          codesObj[DocumentApp.Attribute.BACKGROUND_COLOR] = codesBackgrounds[i][j];
          codesObj[DocumentApp.Attribute.FONT_SIZE] = codesStyles[i][j].getFontSize();
          if(codesStyles[i][j].isBold()) {
            codesObj[DocumentApp.Attribute.BOLD] = true;
          }
          codesTable.getRow(i).getCell(j).setAttributes(codesObj);
      }
    }
  
    createDoc.appendPageBreak();
  
    // 5th page - Claims
    var claimsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Claims');
    
    createDoc.appendParagraph(claimsSheet.getRange(2, 2).getValue()).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    createDoc.appendParagraph("");
  
    var claimsCells = claimsSheet.getRange(4, 2, 29, 4).getValues();
    createDoc.appendListItem(claimsCells.toString()); // ??????
  
    createDoc.appendParagraph("");
  
    // 6th page - Application Events
    var eventsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Application Events');
    
    createDoc.appendParagraph(eventsSheet.getRange(2, 1).getValue()).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    createDoc.appendParagraph("");
    
    var eventsRange = eventsSheet.getRange(4, 1, 8, 2).getDataRegion(SpreadsheetApp.Dimension.ROWS);
    var eventsValues = eventsRange.getDisplayValues();
    var eventsBackgrounds = eventsRange.getBackgrounds();
    var eventsStyles = eventsRange.getTextStyles();
    var eventsTable = createDoc.appendTable(eventsValues);
    //var eventsUrl = eventsSheet.getRange(4, 1, 8, 2).getRichTextValue().getLinkUrl();
  
    eventsTable.setBorderWidth(1);
    
    for(var i = 0; i < eventsTable.getNumRows(); i++) {
      for(var j = 0; j < eventsTable.getRow(i).getNumCells(); j++) {
        var eventsObj = {};
          eventsObj[DocumentApp.Attribute.BACKGROUND_COLOR] = eventsBackgrounds[i][j];
          eventsObj[DocumentApp.Attribute.FONT_SIZE] = eventsStyles[i][j].getFontSize();
          eventsObj[DocumentApp.Attribute.FOREGROUND_COLOR] = eventsStyles[i][j].getForegroundColor();
          if(eventsStyles[i][j].isBold()) {
            eventsObj[DocumentApp.Attribute.BOLD] = true;
          }
          //eventsObj[DocumentApp.Attribute.LINK_URL] = eventsUrl[i][j];
          eventsTable.getRow(i).getCell(j).setAttributes(eventsObj);
      }
    }
  }