function onOpen() {
  // Add a menu with some items, some separators, and a sub-menu.
  DocumentApp.getUi().createMenu('Latex')
      .addItem('Convert to Latex', 'convertToLatex')
      .addToUi();
}

function convertToLatex() {
  var bodyElement = DocumentApp.getActiveDocument().getBody();
  var numChildren = bodyElement.getNumChildren();
  var output = "";
  
  for(var i = 0; i<numChildren; i++){
    var child = bodyElement.getChild(i);
    
    // If paragraph, the element found might be text, image, etc.
    if(child.getType() == "PARAGRAPH"){
      var pNumChildren = child.getNumChildren(); // Number of children in the paragraph
      if(pNumChildren > 0){
        
        for(var k=0; k<pNumChildren; k++){
          var pChild = child.getChild(k);
          
          if(pChild.getType() == "TEXT"){
            output += getTextAsLatex(pChild.editAsText());
          }
          
        }
        
        output += "\n\n";
      }
    }
    
    // If table
    if(child.getType() == "TABLE"){
      output += getTableAsLatex(child);
    }
    
  }
  DocumentApp.getUi().alert(output);
}

function getTextAsLatex(text){
  output = text.getText();
  offset = 0; // Contains the offset to the output string by appending the latex syntax
  indices = text.getTextAttributeIndices();
  
  // Go through all the attribute indices (bold, italic, ..)
  for(i=0; i<indices.length; i++){
    
    // If bold attribute
    if(text.isBold(indices[i])){
      var latexStart = "\\textbf{";
      var latexEnd = "}";
      // Insert latex syntax at start index of attribute
      output = insertStringAtIndex(output, latexStart, indices[i]+offset);
      offset += latexStart.length;
      // Insert latex syntax at end index of attribute
      output = insertStringAtIndex(output, latexEnd, indices[i+1]+offset);
      offset += latexEnd.length;
    }
    
    // If italic attribute
    if(text.isItalic(indices[i])){
      var latexStart = "\\textit{";
      var latexEnd = "}";
      // Insert latex syntax at start index of attribute
      output = insertStringAtIndex(output, latexStart, indices[i]+offset);
      offset += latexStart.length;
      // Insert latex syntax at end index of attribute
      output = insertStringAtIndex(output, latexEnd, indices[i+1]+offset);
      offset += latexEnd.length;
    }
    
    // If underline attribute
    if(text.isUnderline(indices[i])){
      var latexStart = "\\underline{";
      var latexEnd = "}";
      // Insert latex syntax at start index of attribute
      output = insertStringAtIndex(output, latexStart, indices[i]+offset);
      offset += latexStart.length;
      // Insert latex syntax at end index of attribute
      output = insertStringAtIndex(output, latexEnd, indices[i+1]+offset);
      offset += latexEnd.length;
    }
    
  }
  return output;
}
        
function getTableAsLatex(table){

  var output = "";
  var numRows = table.getNumRows();
  
  for(var i=0; i<numRows; i++){
    var row = table.getRow(i);
    var numCells = row.getNumCells();
    
    // Define the table header with number and alignment of columns
    output += "\\begin{tabular}{";  
    for(var i=0; i<numCells; i++){
      output += "l";
      if(i<numCells-1) output += " ";
    }
    output += "}\n";
    
  }
  
  return output;
}

function insertStringAtIndex(text, insertion, index){
  return [text.slice(0, index), insertion, text.slice(index)].join("");
}