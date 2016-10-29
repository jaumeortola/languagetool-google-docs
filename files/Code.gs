/**
 * Server side code
 */
var DIALOG_TITLE = 'Example Dialog';
var SIDEBAR_TITLE = 'LanguageTool proof­reading';

var LT_SERVER = 'https://languagetool.org/api/v2/';
//var LT_SERVER = 'https://www.softcatala.org/languagetool/api/v2/';

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem("Check", 'showSidebar')
    //.addItem("Options", 'showDialog')
    .addToUi();
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

function CheckText(language) {
  language = typeof language !== 'undefined' ? language : 'auto';
  var text = DocumentApp.getActiveDocument().getBody().getText();
  // avoid some bugs
  var cleanText = text.replace(/\n/g,"\n\n").replace(/[ \t]+\n/g, "\n");
  var options = {
    "method": "post",
    "payload": "text=" + encodeURIComponent(cleanText) + "&language=" + language + "&useragent=googledocs"
  };
  var response = UrlFetchApp.fetch(LT_SERVER + "check", options);
  return response.getContentText();
}

function GetLanguages() {
  var response = UrlFetchApp.fetch(LT_SERVER + "languages");
  return response.getContentText();
}

function SelectText(cntxtBefore, cntxtError, cntxtAfter, replacement) {
  if (contextError === "") {
    return "NotFound";
  }
  var contextBefore = escapeRegExp(cntxtBefore);
  var contextError = escapeRegExp(cntxtError);
  var contextAfter = escapeRegExp(cntxtAfter);
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  var rangeBefore;
  var rangeError;
  var rangeAfter;
  var rangeAux;

  if (contextBefore != "") {
    rangeBefore = body.findText(contextBefore);
  }
  if (contextAfter != "") {
    rangeAfter = body.findText(contextAfter);
  }
  if (rangeBefore != null) {
    // try to find the error after rangeBefore 
    rangeError = body.findText(contextError, rangeBefore);    
  } else if (rangeAfter != null) {
      // try to find from the start of the paragraph containing rangeAfter
      rangeError = rangeAfter.getElement().asText().findText(contextError);
      while (rangeAfter.getStartOffset() - rangeError.getEndOffsetInclusive() > 2 &&
        rangeAfter.getStartOffset() > rangeError.getEndOffsetInclusive()) {
          rangeError = body.findText(contextError, rangeError);
        }
  }
  // try to find the paragraph containing the "short" context. 
  if (rangeError == null) {
    var shortContext = escapeRegExp(getShortContext(cntxtBefore, cntxtError, cntxtAfter, 5));
    rangeAux = body.findText(shortContext);
    if (rangeAux != null) {
      rangeError = rangeAux.getElement().asText().findText(contextError);
    }
  } 
  // the error is at the sentence end
  if (rangeError == null) {
    var shortContext = escapeRegExp(getShortContext(cntxtBefore, cntxtError, "", 5));
    rangeAux = body.findText(shortContext);
    if (rangeAux != null) {
      rangeError = rangeAux.getElement().asText().findText(contextError);
    }
  } 
  // the error is at the sentence start
  if (rangeError == null) {
    var shortContext = escapeRegExp(getShortContext("", cntxtError, cntxtAfter, 5));
    rangeAux = body.findText(shortContext);
    if (rangeAux != null) {
      rangeError = rangeAux.getElement().asText().findText(contextError);
    }
  } 
  // Catch-all case. Try to find error anywhere.
  if (rangeError == null) {
    rangeError = body.findText(contextError);
  }
 
  if (rangeError == null || (rangeAfter != null && rangeError.getEndOffsetInclusive() > rangeAfter.getStartOffset())) {
    return "NotFound";
  }

  var rangeBuilder = doc.newRange();
  rangeBuilder.addElement(rangeError.getElement(), rangeError.getStartOffset(), rangeError.getEndOffsetInclusive());
  doc.setSelection(rangeBuilder.build());

  if (replacement && replacement.length > 0) {
    var startOffset = rangeError.getStartOffset()
    var endOffset = rangeError.getEndOffsetInclusive()
    rangeError.getElement().asText().deleteText(startOffset, endOffset)
    rangeError.getElement().asText().insertText(startOffset, replacement)
    return "Replaced";
  }
  return "Selected";
}


function getShortContext(contextBefore, contextError, contextAfter,len) {
  var shortCntxt = "";
  if (contextBefore.length >= len) {
    shortCntxt = contextBefore.slice(-len);
  } else {
    shortCntxt = contextBefore;
  }
  shortCntxt = shortCntxt + contextError;
  if (contextAfter.length >= len) {
    shortCntxt = shortCntxt + contextAfter.substring(0,len);
  } else {
    shortCntxt = shortCntxt + contextAfter;
  }
  return shortCntxt;
}

function escapeRegExp(str) {
  return str.replace(/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|]/g, "\\$&");
}

/**
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle(SIDEBAR_TITLE);
  DocumentApp.getUi().showSidebar(ui);
}

/**
 * Opens a dialog. The dialog structure is described in the Dialog.html
 * project file.
 */
function showDialog() {
  var ui = HtmlService.createTemplateFromFile('Dialog')
    .evaluate()
    .setWidth(400)
    .setHeight(190);
  DocumentApp.getUi().showModalDialog(ui, DIALOG_TITLE);
}
