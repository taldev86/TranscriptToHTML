/* What should the add-on do when a document is opened */
function onOpen() {
  DocumentApp.getUi()
  .createAddonMenu() // Add a new option in the Google Docs Add-ons Menu
  .addItem("Run", "CONVERT")
  .addToUi();  // Run the showSidebar function when someone clicks the menu
}

function CONVERT() {
  var body = DocumentApp.getActiveDocument().getBody();
  var numChildren = body.getNumChildren();
  var output_final = [];

  // Walk through all the child elements of the body.
  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    output_final.push(processItem(child));
  }

  var html = output_final.join('\n');
  emailHtml(html);
  //return ContentService.createTextOutput(html);
}



function processItem(item) {
  var output = [];
  
  //if reading text inside a paragraph element
  if (item.getType() == DocumentApp.ElementType.TEXT) {
    //process the text! next function...
    processText(item, output);
    
  //if processing paragraph element itself, send child text element back into this function to be processed
  } else if (item.getType() == DocumentApp.ElementType.PARAGRAPH) {
    //if child elements exist (hopefully just 1 TEXT element) put it back into this function for text processing (condition above)!
    if (item.getNumChildren) {
      var numChildren = item.getNumChildren();
      for (var i = 0; i < numChildren; i++) {
        var child = item.getChild(i);
        output.push(processItem(child));
      }
    }

  //if element doesn't qualify, skip it and leave processing
  } else {
    return "";
  }
  
  //stringify array, to be added to final output array
  return output.join('');
}



function processText(item, output) {
  var text = item.getText().trim();
  var indices = item.getTextAttributeIndices();
  
  //if empty line, leave
  if (text == '') {
    return;
    
  //Sound Effects: line wrapped in []
  } else if (text.substring(0,1) == '[' && text.slice(-1) == ']') {
    output.push('<p class="sound prerecorded">');
    output.push(text);
    output.push('</p>');
    
  } else {
    
    //make sure there is content in paragraph
    if (text.length > 1) {
    
      //Check attributes at start of line to determine type
      var paraAtts = item.getAttributes(indices[0]+1);
      
      //if Speaker line, split up text to wrap speaker in span and add array
      if (paraAtts.BOLD) {
        //get speaker and word as two units, remove colon and extra spaces (will add back in programatically)
        var speaker = text.substring(0, indices[1]).trim();
        speaker = speaker.slice(-1) == ":" ? speaker.slice(0, -1) : speaker;
        var words = text.substring(indices[1]).trim();
        words = words.substring(0,1) == ":" ? words.substring(1, words.length) : words;
        //make markup
        output.push('<span class="speaker">', speaker, ': ', '</span>', words);
        
        //otherwise it's continues speach, just add text
      } else {
        output.push(text);
      }
      
      //open paragraph with class indicating narrator or pre-recorded
      //Pre-recorded: line starts italic
      if (paraAtts.ITALIC) {
        output.unshift('<p class="prerecorded">');
        
        //Narrator, line starts not italic
      } else {
        output.unshift('<p class="narrator">');
      }
      
      //close paragraph
      output.push('</p>');
      
      return output.join('');
    }
    
  }
}



function emailHtml(html) {
  var attachments = [];
  var name = DocumentApp.getActiveDocument().getName()+".html";
  attachments.push({"fileName":name, "mimeType": "text/html", "content": html});
  MailApp.sendEmail({
     to: Session.getActiveUser().getEmail(),
     subject: name,
     htmlBody: html,
     attachments: attachments
   });
}
