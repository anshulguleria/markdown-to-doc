/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
    DocumentApp.getUi().createAddonMenu()
        .addItem('Format Selection', 'showPopupHtml')
        .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
    onOpen(e);
}

function showPopupHtml() {
    var converter = getConverter();

    var selection = DocumentApp.getActiveDocument()
        .getSelection();
    if(selection) {
        var elements = selection.getRangeElements();
        var markdown = "";
        for(var i = 0 ; i < elements.length; i++) {
            var element = elements[i].getElement();
            Logger.log(element.getType());
            markdown += element.asText().getText();
            if(element.getType() !== DocumentApp.ElementType.TEXT) {
                markdown += "\n";
            }
        }

        Logger.log(markdown);

        var htmlString = converter.makeHtml(markdown);
        var htmlTemplate = HtmlService.createTemplateFromFile("htmlpopup");
        htmlTemplate.html_text = htmlString;

        var html = HtmlService.createHtmlOutput(htmlTemplate.evaluate().getContent())
            .setWidth(600)
            .setHeight(400);

        DocumentApp.getUi()
            .showModalDialog(html, "Markdown Preview");
    } else {
        DocumentApp.getUi().alert("No selection made");
    }
}

function deleteSelection () {
    var doc = DocumentApp.getActiveDocument();
    var selection = doc
        .getSelection();

    var newPosition = null;

    if(selection) {
        var elements = selection.getRangeElements();
        for(var i = 0; i < elements.length - 1; i++) {
            var element = elements[i].getElement();
            element.removeFromParent();
        }

        // Dont' delete last element only clear it
        var element = elements[i].getElement();
        element.asText().setText("");
        newPosition = doc.newPosition(element, 0);
    }

    return newPosition;
}

function replaceWithRichText (text) {
    Logger.log("got text as: " + text);
    // Remove selected text
    //var newPosition = deleteSelection();

    //  var doc = DocumentApp.getActiveDocument();
    //  // Add our text
    //  if(newPosition) {
    //    var didInsert = newPosition.insertText(text);
    //    if(!didInsert) {
    //      DocumentApp.getUi().alert("Can't insert at the selected position.");
    //    }
    //  } else {
    //    DocumentApp.getUi().alert("Can't find cursor.");
    //  }
    //
    return true;
}

var cachedConverter = null;

function getConverter () {
    if(!cachedConverter) {
        cachedConverter = new showdown.Converter();
    }

    return cachedConverter;

    //text      = '#hello, markdown!',
    //html      = converter.makeHtml(text);
}
