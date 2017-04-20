function onOpen(e){
  DocumentApp.getUi() 
  .createAddonMenu()
  .addItem('Highlight (Default yellow color)', 'highlightVerbs')
  .addItem("Highlight (Specific color)", 'openSettings')
  .addToUi();
}  

function highlightVerbs(color) {
  
  if (color == null) color = "#FCFC00";
  
  var numOfVerbs = 0;
  
  var body = DocumentApp.getActiveDocument().getBody();
  
  var dictionary = [
    " are ",
    " can ",
    " is ",
    " am ",
    " was ",
    " were ",
    " might ",
    " must ",
    " may ",
    " had ",
    " has ",
    " have ",
    " having ",
    " be ",
    " been ",
    " should ",
    " could ",
    " would ",
    " will ",
    " did ",
    " do ",
    " done ",
    " does ",
    " doing "
  ];
  
  for(var i = 0; i < dictionary.length; i++){

    var foundElement = body.findText(dictionary[i]);
  
  

    while (foundElement != null) {
        // Get the text object from the element
      var foundText = foundElement.getElement().asText();
      
      // Where in the Element is the found text?
      var start = foundElement.getStartOffset() + 1;
      var end = foundElement.getEndOffsetInclusive() - 1;
      
      // Change the background color to yellow
      foundText.setBackgroundColor(start, end, color);
      
      // Find the next match
      foundElement = body.findText(dictionary[i], foundElement);
      
      numOfVerbs++;
    }
  }
  
  var leftApos = String.fromCharCode(8220);
  var rightApos = String.fromCharCode(8221);
  
  var firstElement = body.findText(leftApos);
  var secondElement = body.findText(rightApos, firstElement);
  
  while(secondElement != null || firstElement != null){
    var firstText = firstElement.getElement().asText();
    var secondText = secondElement.getElement().asText();
    
    var first = firstElement.getStartOffset();
    var second = secondElement.getEndOffsetInclusive();
    
    firstText.setBackgroundColor(first, second, "#ffffff");
    
    firstElement = body.findText(leftApos, secondElement);
    secondElement = body.findText(rightApos, firstElement);
    if(firstElement == null || secondElement == null){break;}
  }
  
  showPopup(numOfVerbs);
  
}

function showPopup(number){
  var ui = DocumentApp.getUi();
  
  ui.alert("Number of To-Be Verbs: " + number + "\n(Includes To Be Verbs in quotes)");
}

function openSettings(){
  var ui = DocumentApp.getUi();
  
  var result = ui.prompt("Highlight in a specific color","Enter a color in CSS hex", ui.ButtonSet.OK_CANCEL)
  
  var button = result.getSelectedButton();
  var newColor = result.getResponseText();
  Logger.log(newColor);
  
  if(button === ui.Button.OK){
    if(newColor === ''){
      ui.alert("Please enter a color!");
    } else {
      highlightVerbs(newColor)
      Logger.log("Changed color to: " + newColor);
    }
  }
}
