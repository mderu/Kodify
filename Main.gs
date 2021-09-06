// Based on StackOverflow-Light style
// https://github.com/highlightjs/highlight.js/blob/main/src/styles/stackoverflow-light.css
styles = {
  "subst": (editableText, start, end) => editableText.setForegroundColor(start, end, "#2f3337"),
  "comment": (editableText, start, end) => editableText.setForegroundColor(start, end, "#656e77"),
  "keyword":      (editableText, start, end) => styles["section"](editableText, start, end),
  "selector-tag": (editableText, start, end) => styles["section"](editableText, start, end),
  "meta":         (editableText, start, end) => styles["section"](editableText, start, end),
  "doctag":       (editableText, start, end) => styles["section"](editableText, start, end),
  "section":      (editableText, start, end) => editableText.setForegroundColor(start, end, "#015692"),
  "attr": (editableText, start, end) => editableText.setForegroundColor(start, end, "#015692"),
  "attribute": (editableText, start, end) => editableText.setForegroundColor(start, end, "#803378"),
  "name":         (editableText, start, end) => styles["template-tag"](editableText, start, end),
  "type":         (editableText, start, end) => styles["template-tag"](editableText, start, end),
  "number":       (editableText, start, end) => styles["template-tag"](editableText, start, end),
  "selector-id":  (editableText, start, end) => styles["template-tag"](editableText, start, end),
  "quote":        (editableText, start, end) => styles["template-tag"](editableText, start, end),
  "template-tag": (editableText, start, end) => editableText.setForegroundColor(start, end, "#b75501"),
  "selector-class": (editableText, start, end) => editableText.setForegroundColor(start, end, "#015692"),
  "string":            (editableText, start, end) => styles["selector-attr"](editableText, start, end),
  "regexp":            (editableText, start, end) => styles["selector-attr"](editableText, start, end),
  "symbol":            (editableText, start, end) => styles["selector-attr"](editableText, start, end),
  "variable":          (editableText, start, end) => styles["selector-attr"](editableText, start, end),
  "template-variable": (editableText, start, end) => styles["selector-attr"](editableText, start, end),
  "link":              (editableText, start, end) => styles["selector-attr"](editableText, start, end),
  "selector-attr":     (editableText, start, end) => editableText.setForegroundColor(start, end, "#54790d"),
  "meta":            (editableText, start, end) => styles["selector-pseudo"](editableText, start, end),
  "selector-pseudo": (editableText, start, end) => editableText.setForegroundColor(start, end, "#015692"),
  "built_in": (editableText, start, end) => styles["literal"](editableText, start, end),
  "title":    (editableText, start, end) => styles["literal"](editableText, start, end),
  "literal":  (editableText, start, end) => editableText.setForegroundColor(start, end, "#b75501"),
  "bullet": (editableText, start, end) => styles["code"](editableText, start, end),
  "code":   (editableText, start, end) => editableText.setForegroundColor(start, end, "#535a60"),
  // meta & string?? => editableText.setForegroundColor(start, end, "#54790d"),
  "deletion": (editableText, start, end) => editableText.setForegroundColor(start, end, "#c02d2e"),
  "addition": (editableText, start, end) => editableText.setForegroundColor(start, end, "#2f6f44"),
  "emphasis": (editableText, start, end) => editableText.setItalic(start, end, true),
  "strong": (editableText, start, end) => editableText.setBold(start, end, true),
  // Purposely set underline to false (it is never underlined) to return a Text object.
  "formula": (editableText, start, end) => editableText.setUnderline(start, end, false),
  "operator": (editableText, start, end) => editableText.setUnderline(start, end, false),
  "params": (editableText, start, end) => editableText.setUnderline(start, end, false),
  "property": (editableText, start, end) => editableText.setUnderline(start, end, false),
  "punctuation": (editableText, start, end) => editableText.setUnderline(start, end, false),
  "tag": (editableText, start, end) => editableText.setUnderline(start, end, false),
  // Function was missing
  "function": (editableText, start, end) => editableText.setUnderline(start, end, false),
  // Add a default, so we can clear any highlighting on strings that users may have accidentally made
  // that don't actually have a highlight color, or that was copy/pasted from another source that had
  // a different font or highlighting (e.g., Intellij)
  "default": (editableText, start, end) => {editableText.setFontFamily(start, end, "Consolas"); editableText.setForegroundColor(start, end, "#000000");}
}

var tableStyles = {
  "bad":  {[DocumentApp.Attribute.BACKGROUND_COLOR]: "#fff9f9", [DocumentApp.Attribute.BORDER_COLOR]: "#f0ccc5", [DocumentApp.Attribute.FONT_FAMILY]: "Consolas" },
  "good": {[DocumentApp.Attribute.BACKGROUND_COLOR]: "#f9fff9", [DocumentApp.Attribute.BORDER_COLOR]: "#cef3c9", [DocumentApp.Attribute.FONT_FAMILY]: "Consolas" },
  "":     {[DocumentApp.Attribute.BACKGROUND_COLOR]: "#fbfbfb", [DocumentApp.Attribute.BORDER_COLOR]: "#e0e0e0", [DocumentApp.Attribute.FONT_FAMILY]: "Consolas" },
}

function main() {
  cleanRogueLineBreaks()
  handleInlineCode()
  turnFencesIntoTables()
  highlightTables()
}

function onOpen() {
  var ui = DocumentApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu("Kodify").addItem("Apply Style", "main").addToUi()
}

// Sometimes, paragraphs can have new lines created by "\r" characters inside of them.
// It sounds dumb because it is dumb. We need to properly split these paragraphs so our
// formatting works.
// This may break some backwards compat with documents coming from Microsoft Word, but that is
// a risk that I am willing to take.
function cleanRogueLineBreaks() {
  var body = DocumentApp.getActiveDocument().getBody();
  var childNum = 0;
  while (childNum < body.getNumChildren()) {
    var child = body.getChild(childNum)
    // Ingore text in tables.
    if (child.getType() == DocumentApp.ElementType.TABLE) {childNum++; continue;}
    if (child.getType() == DocumentApp.ElementType.TABLE_ROW) {childNum++; continue;}
    if (child.getType() == DocumentApp.ElementType.TABLE_CELL) {childNum++; continue;}

    if (child.asText() !== null) {
      lines = child.asText().getText().split("\r")
      if (lines.length > 1) {
        console.log("Found a rogue newline!:\n" + child.asText().getText())
        child.asText().setText(lines[0])
        for (var i = 1; i < lines.length; i++) {
          body.insertParagraph(childNum + i, lines[i])
            .setAttributes(child.getAttributes())
        }
      }
    }
    childNum++;
  }
}

// Injects startRange and stopRange into the emittedNodes from the parsed tree.
function injectRanges(emittedNode, cursor) {
  emittedNode.startRange = cursor;
  length = 0;
  for (const childNode of emittedNode.children) {
    if (typeof childNode === "string") {
      length += childNode.length
    }
    else {
      length += injectRanges(childNode, cursor + length)
    }
  }
  emittedNode.endRange = cursor + length;
  return length;
}

// Applies the specified style dictionary to the given text
function applyStyling(emittedNode, editableText, style) {
  for (const childNode of emittedNode.children) {
    if (typeof childNode === "string") { continue; }
    var kind = childNode.kind.split(".")[0]
    // Sometimes we get kinds that are kindA.kindB, where kindB doesn't seem to be
    // in the CSS code. Just going to default it to kindA.
    if (kind != childNode.kind) {
      console.log("Changing " + childNode.kind + " for " + kind)
    }
    // If somehow it comes back with a kind we don't have, just default instead of
    // erroring out.
    if (! (kind in style)) {
      console.log("Unable to find `" + kind + "`, defaulting.");
      kind = "default";
    }
    // Subtract 1 because endRange is exclusive, but formatting rules are inclusive.
    style[childNode.kind](editableText, childNode.startRange, childNode.endRange - 1)
    applyStyling(childNode, editableText, style);
  }
}

// Converts inlined code to code style
function handleInlineCode() {
  var body = DocumentApp.getActiveDocument().getBody();
  var inlineStyle = {}
  inlineStyle[DocumentApp.Attribute.FONT_FAMILY] = "Consolas"
  inlineStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = "#2f9b00"
  lastRange = DocumentApp.getActiveDocument().newRange().build().getRangeElements()[0];
  while (true) {
    var inlineRangeElement = body.findText("`[^`]+`", lastRange)
    if (inlineRangeElement == null) {break;}
    inlineRangeElement.getElement().setAttributes(inlineRangeElement.getStartOffset(), inlineRangeElement.getEndOffsetInclusive(), inlineStyle)
    inlineRangeElement.getElement().deleteText(inlineRangeElement.getEndOffsetInclusive(), inlineRangeElement.getEndOffsetInclusive())
    inlineRangeElement.getElement().deleteText(inlineRangeElement.getStartOffset(), inlineRangeElement.getStartOffset())
    lastRange = inlineRangeElement;
  }
}

function turnFencesIntoTables() {
  var body = DocumentApp.getActiveDocument().getBody();

  var childNum = 0;
  var fenceStart = null;
  var fenceStartIndex = 0
  var allTextBetweenFences = []
  while (childNum < body.getNumChildren()) {
    var child = body.getChild(childNum);
    if (child.getType() == DocumentApp.ElementType.PAGE_BREAK
        || child.getType() == DocumentApp.ElementType.TABLE
        || child.getType() == DocumentApp.ElementType.LIST_ITEM
        || child.getType() == DocumentApp.ElementType.FOOTER_SECTION
        || child.getType() == DocumentApp.ElementType.HEADER_SECTION
        || child.getType() == DocumentApp.ElementType.TABLE_OF_CONTENTS
        ) {
      childNum++;
      fenceStart = null;
      continue;
    }
    if (child.getType() == DocumentApp.ElementType.PARAGRAPH 
        && (child.asParagraph().getHeading() != DocumentApp.ParagraphHeading.NORMAL ||
        child.asParagraph().findElement(DocumentApp.ElementType.PAGE_BREAK) != null)
        ) {
      childNum++;
      fenceStart = null;
      continue;
    }
    // Attempt to set the fence start
    if (fenceStart == null) {
      fenceStart = child.asText().findText("^```([Ss]tyle:(bad|good))?(\\r?\\n.*)?$")
      fenceStartIndex = childNum;
    }
    else {
      var fenceEnd = child.asText().findText("^```");
      if (fenceEnd != null) {
        var newTable = body.insertTable(fenceStartIndex, new Array(new Array(allTextBetweenFences.join("\n"))));
        var style = fenceStart.getElement().asText().findText(":good") ? "good" :
            fenceStart.getElement().asText().findText(":bad") ? "bad" : "";
        newTable.setAttributes(tableStyles[style]);
        var parentIndent = fenceStart.getElement().getParent().getIndentStart()
        // Ugly hack to indent table
        if (parentIndent > 0) {
          body.removeChild(newTable)
          insertTableWithIndentation(
            body,
            DocumentApp.getActiveDocument(),
            newTable,
            fenceStartIndex,
            parentIndent)
        }
        newTable.getCell(0, 0).setAttributes(tableStyles[style]);
        childNum += 1;
        fenceStartIndex += 1;
        for (var i = childNum; i >= fenceStartIndex; i--) {
          // You are not allowed to remove the last paragraph in a body section, so clear out the line instead.
          if (body.getChild(fenceStartIndex).isAtDocumentEnd()) {
            body.getChild(i).asText().editAsText().setText("");
            continue;
          }
          body.getChild(fenceStartIndex).removeFromParent()
          // Need to reduce the childNum here to make sure the next fence can be picked up properly after lines are removed.
          childNum--;
        }
        fenceStart = null;
        allTextBetweenFences = [];
      } else {
        allTextBetweenFences.push(child.asText().getText())
      }
    }
    childNum++;
  }
}

function highlightElement(element) {
  var text = element.getText();
  if (text.length == 0) {
    return;
  }
  var emitter = hljs.highlightAuto(text)._emitter;
  injectRanges(emitter.rootNode, 0);

  // The user may have changed the code in-line, and picked up a foreground color of a word
  // that was not the default color. When symbols that do not get a highlight color are formatted,
  // they will maintain this highlight color. To prevent this, we apply the default style to
  // the entire element.
  styles["default"](element.editAsText(), 0, text.length - 1)

  applyStyling(emitter.rootNode, element.editAsText(), styles);
}

function highlightTables() {
  var tables = DocumentApp.getActiveDocument().getBody().getTables()
  for (var table of DocumentApp.getActiveDocument().getBody().getTables()) {
    if (table.getNumRows() != 1) {continue;}
    if (table.getRow(0).getNumCells() != 1) {continue;}
    highlightElement(table.getRow(0).getCell(0));
  }
}

// https://stackoverflow.com/a/51474209/6876989
var insertTableWithIndentation = function(body, 
                                        document, 
                                        originalTable, 
                                        elementIndex, 
                                        indentWidth) {

  var docWidth = (document.getPageWidth() - 
                  document.getMarginLeft() - 
                  document.getMarginRight());

  var table = body.insertTable(elementIndex);
  var attrs = table.getAttributes();
  attrs['BORDER_WIDTH'] = 0;
  table.setAttributes(attrs);

  var row = table.appendTableRow();
  var column1Width = indentWidth;
  row.appendTableCell().setWidth(column1Width);

  var cell = row.appendTableCell();
  var column2Width = docWidth - column1Width;
  cell.setPaddingTop(0)
      .setPaddingBottom(0)
      .setPaddingLeft(0)
      .setPaddingRight(0)
      .setWidth(column2Width);

  cell.insertTable(0, originalTable);
}
