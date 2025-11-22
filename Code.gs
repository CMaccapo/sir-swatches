/** @OnlyCurrentDoc */
function isValidHexColor(str) {
  return /^#([0-9A-F]{3}|[0-9A-F]{6})$/i.test(str);
}

function colorTerm(term, color, thisBtnId) {
  const body = DocumentApp.getActiveDocument().getBody();
  let found = body.findText(term);
  while (found !== null) {
    const el = found.getElement();
    const start = found.getStartOffset();
    const end = found.getEndOffsetInclusive();

    if (thisBtnId == "highlightBtn"){
      el.setBackgroundColor(start, end, color);
    }
    if (thisBtnId == "colorizeBtn"){
      el.setForegroundColor(start, end, color);
    }
    found = body.findText(term, found);
  }
}

function swapColors(oldColor, newColor, thisBtnId) {
  if (typeof oldColor === "string" && typeof newColor === "string" && oldColor.trim() && newColor.trim()) {
    if (isValidHexColor) {
      oldColor = oldColor.toLowerCase();
      if (oldColor == null){
        oldColor = "#000000";
      }
      newColor = newColor.toLowerCase();
      const body = DocumentApp.getActiveDocument().getBody();
      const textElements = getAllTextElements(body);

      textElements.forEach(el => {
        const text = el.getText();
        if (thisBtnId == "highlightBtn"){
          for (let i = 0; i < text.length && typeof el.getBackgroundColor(i) === "string"; i++) {
            if (el.getBackgroundColor(i).toLowerCase() === oldColor) {  //if highlighted section is input color
              el.setBackgroundColor(i, i, newColor);
            }
          }
        }
        if (thisBtnId == "colorizeBtn"){
          for (let i = 0; i < text.length && typeof el.getForegroundColor(i) === "string"; i++) {
            if (el.getForegroundColor(i).toLowerCase() === oldColor) {  //if highlighted section is input color
              el.setForegroundColor(i, i, newColor);
            }
          }
        }
      });
    }
  }
}

function getAllTextElements(container) {
  const elements = [];
  const numChildren = container.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const element = container.getChild(i);
    const type = element.getType();
    if (type === DocumentApp.ElementType.TEXT) {
      elements.push(element.asText());
    } else if (element.getNumChildren) {
      elements.push(...getAllTextElements(element));
    }
  }
  return elements;
}

function changeColors(thisBtnId, bgColor, fgColor){
  if (thisBtnId.toLowerCase().includes("highlight")){
    if (bgColor !== null) {
      return bgColor;
    }
  }
  if (thisBtnId.toLowerCase().includes("color")){
    if (fgColor !== null) {
      return fgColor;
    }
    else {
      return "#000000";
    }
  }
}

function getColorFromSelection(thisBtnId) {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();

  if (selection) {
    const elements = selection.getRangeElements();

    for (let i = 0; i < elements.length; i++) {
      const rangeElement = elements[i];

      if (rangeElement.getElement().editAsText) {
        const text = rangeElement.getElement().editAsText();
        const start = rangeElement.getStartOffset();
        
        return changeColors(thisBtnId, text.getBackgroundColor(start), text.getForegroundColor(start));
      }
    }
  } 
  else {
    // No selection — try to get the cursor position
    const cursor = doc.getCursor();
    if (cursor) {
      const element = cursor.getElement();
      if (element.editAsText) {
        const text = element.editAsText();
        const offset = cursor.getOffset();

        if (offset !== -1) {
          return changeColors(thisBtnId, text.getBackgroundColor(offset), text.getForegroundColor(offset));
        }
      }
    }
  }

  return null; // no highlight color found
}

function testPrint(toPrint) {
  console.log("AAAAA: "+toPrint);
}

function showDialog() {
  const html = HtmlService.createHtmlOutputFromFile("Dialog")
    .setWidth(250)
    .setHeight(290);
  DocumentApp.getUi().showModelessDialog(html, "SirSwatches");
}
function onOpen() {
  DocumentApp.getUi()
    .createMenu("🖍️")
    .addItem("Open", "showDialog")
    .addToUi();
}

function main() {
  onOpen()
}
main()
