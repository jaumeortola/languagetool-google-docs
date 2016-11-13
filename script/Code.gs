/**
 * Server side code
 */
var DIALOG_TITLE = 'LanuageTool Options';
var SIDEBAR_TITLE = 'LanguageTool proof­reading';

var LT_SERVER = 'https://languagetool.org/api/v2/';
//var LT_SERVER = 'https://www.softcatala.org/languagetool/api/v2/';
var MAX_CHAR_LENGTH = 20000;
var MIN_CHAR_LENGTH = 20;

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem("Check", 'showSidebar')
    .addItem("Options", 'showDialog')
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

function getUserProperties() {
  var userProp = PropertiesService.getUserProperties();
  if (userProp.getProperty("LT_SERVER") == null) {
    userProp.setProperty("VARIANT_EN", "en-US");
    userProp.setProperty("VARIANT_DE", "de-DE");
    userProp.setProperty("VARIANT_PT", "pt-PT");
    userProp.setProperty("VARIANT_CA", "ca-ES");
    userProp.setProperty("LT_SERVER", LT_SERVER);
  }
  if (userProp.getProperty("PERSONAL_DICT") == null) {
    userProp.setProperty("PERSONAL_DICT", "");
  }
  return userProp.getProperties();
}


function getDocumentProperties() {
  var docProp = PropertiesService.getDocumentProperties();
  if (docProp.getProperty("DISABLED_RULES") == null) {
    docProp.setProperty("DISABLED_RULES", "");
  }
  return docProp.getProperties();
}

function CheckText(language) {
  language = typeof language !== 'undefined' ? language : 'auto';
  var text = DocumentApp.getActiveDocument().getBody().getText();
  textTotalLength = text.length;
  var selectedText = getSelectedText();
  if (selectedText && selectedText.length > MIN_CHAR_LENGTH && selectedText.length < MAX_CHAR_LENGTH) {
    text = selectedText;
  } else if (textTotalLength > MAX_CHAR_LENGTH) {
    throw "The document is too long. Select some text between " + MIN_CHAR_LENGTH + " and " + MAX_CHAR_LENGTH + " characters."
  }
  var textCheckedLength = text.length;
  // avoid some bugs
  var cleanText = text.replace(/\n/g, "\n\n").replace(/[ \t]+\n/g, "\n");
  var data = "text=" + encodeURIComponent(cleanText) + "&language=" + language + "&useragent=googledocs";
  var userProp = PropertiesService.getUserProperties();
  if (language == "auto") {
    var preferredVariants = userProp.getProperty("VARIANT_EN") + "," + userProp.getProperty("VARIANT_DE") + "," + userProp.getProperty("VARIANT_PT") + "," + userProp.getProperty("VARIANT_CA");
    data += "&preferredVariants=" + preferredVariants;
  }
  var options = {
    "method": "post",
    "payload": data
  };
  try {
    var response = UrlFetchApp.fetch(getUserProperties().LT_SERVER + "check", options);
  } catch (err) {
    throw 'Error: Cannot conect to the server ' + getUserProperties().LT_SERVER;
  }

  var responseJson = JSON.parse(response.getContentText());
  if (textCheckedLength > textTotalLength) {
    textCheckedLength = textTotalLength
  }
  responseJson.extrainfo = "Checked " + textCheckedLength + " of " + textTotalLength + " characters."
  responseJson.personaldict = userProp.getProperty("PERSONAL_DICT");
  responseJson.disabledrules = getDocumentProperties().DISABLED_RULES;
  responseJson.ltserver = userProp.getProperty("LT_SERVER");

  return JSON.stringify(responseJson);
}

function getSelectedText() {
  var selectedText = "";
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      if (elements[i].isPartial()) {
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();

        selectedText += element.getText().substring(startIndex, endIndex + 1) + "\n";
      } else {
        var element = elements[i].getElement();
        // Only translate elements that can be edited as text; skip images and
        // other non-text elements.
        if (element.editAsText) {
          var elementText = element.asText().getText();
          // This check is necessary to exclude images, which return a blank
          // text element.
          if (elementText != '') {
            selectedText += elementText + "\n";
          }
        }
      }
    }
  }
  return selectedText;
}


function GetLanguages() {
  try {
    var response = UrlFetchApp.fetch(getUserProperties().LT_SERVER + "languages");
  } catch (err) {
    throw 'Error: Cannot conect to the server ' + getUserProperties().LT_SERVER;
  }
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

function getShortContext(contextBefore, contextError, contextAfter, len) {
  var shortCntxt = "";
  if (contextBefore.length >= len) {
    shortCntxt = contextBefore.slice(-len);
  } else {
    shortCntxt = contextBefore;
  }
  shortCntxt = shortCntxt + contextError;
  if (contextAfter.length >= len) {
    shortCntxt = shortCntxt + contextAfter.substring(0, len);
  } else {
    shortCntxt = shortCntxt + contextAfter;
  }
  return shortCntxt;
}

function escapeRegExp(str) {
  return str.replace(/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|]/g, "\\$&").replace(/ /g, "\\s");
}


function processForm(formObject) {
  PropertiesService.getUserProperties().setProperty("LT_SERVER", formObject.lt_server);
  PropertiesService.getUserProperties().setProperty("VARIANT_EN", formObject.variant_en);
  PropertiesService.getUserProperties().setProperty("VARIANT_DE", formObject.variant_de);
  PropertiesService.getUserProperties().setProperty("VARIANT_PT", formObject.variant_pt);
  PropertiesService.getUserProperties().setProperty("VARIANT_CA", formObject.variant_ca);
  PropertiesService.getUserProperties().setProperty("DISABLED_RULES", formObject.disabled_rules);
  PropertiesService.getUserProperties().setProperty("ENABLED_RULES", formObject.enabled_rules);
  PropertiesService.getUserProperties().setProperty("PERSONAL_DICT", formObject.personal_dict);
}

function addToDict(word) {
  var userProp = PropertiesService.getUserProperties();
  var personaldict = userProp.getProperty("PERSONAL_DICT");
  if (personaldict.length > 0) {
    personaldict += ",";
  }
  userProp.setProperty("PERSONAL_DICT", personaldict + word);
}

function addDisabledRule(ruleId) {
  var docProp = PropertiesService.getDocumentProperties();
  var disabledrules = docProp.getProperty("DISABLED_RULES");
  if (disabledrules.length == 0) {
    disabledrules += ",";
  }
  docProp.setProperty("DISABLED_RULES", disabledrules + ruleId + ",");
}

function removeFromDisabledRules(ruleId) {
  var docProp = PropertiesService.getDocumentProperties();
  var disabledrules = docProp.getProperty("DISABLED_RULES");
  var disabledrules = disabledrules.replace(new RegExp("(,|^)" + ruleId + "(,|$)", 'g'), ",");
  docProp.setProperty("DISABLED_RULES", disabledrules);
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
    .setHeight(500);
  DocumentApp.getUi().showModalDialog(ui, DIALOG_TITLE);
}
