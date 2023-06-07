function onFileSelected(e) {
  const items = e.drive.selectedItems;
  const item = items[0];
  if(item.mimeType !== "message/rfc822")
    return onHomepage();
  const card = createCard(getEmlContents(item.id)); // analyse the first of the selected items
  return [card.build()];
}

function onHomepage() {
  return [createCard('Please select an .EML file').build()];
}

function getFileContents(id) {
  try {
    const file = DriveApp.getFileById(id);
    return file.getBlob().getDataAsString();
  } catch (error) {
    throw new Error("Failed to get file contents: " + error.message);
  }
}

function getEmlContents(id) {
  const emlContent = getFileContents(id);
  return analyseEmlContent(emlContent);
}

function analyseEmlContent(emlContent) {
  const emailData = convertEMLToJSON(emlContent);
  console.log(emailData);
  const emailPlainText = emailData.find(x => x['Content-Type']?.indexOf('text/plain') === 0 && (x['Content-Disposition'] ? !x['Content-Disposition'].includes("attachment") : true));
  const emailHtml = emailData.find(x => x['Content-Type']?.indexOf('text/html') === 0 && !x['Content-Disposition']?.includes("attachment"));
  const content = emailPlainText ? emailPlainText.Content : emailHtml ? stripHtmlTags(emailHtml.Content) : '';
  const attachments = emailData.filter(x => x['Content-Disposition']?.indexOf("attachment") === 0); // counts both inline and attachment types, but not text and html inline
  return `Subject: ${emailData[0].Subject}\nFrom: ${emailData[0].From}\nDate: ${emailData[0].Date}\nAttachments: ${attachments.length}\n\nBody: \n\n${content}`;
}

function convertEMLToJSON(text) {
  const result = [];
  const boundaries = findEmlBoundaries(text);
  const parts = boundaries.length > 0 ?
    text.split(new RegExp(boundaries.map(x => x + "\r?\n?").join('|'))).filter(x => x) : [text];
    // if there are no boundaries, just return the whole text as one part
  
  for(let p=0; p<parts.length; p++) {
    const lines = parts[p].split('\n');
    const jsonObj = {};
    for(let i=0; i<lines.length; i++) {
        if (lines[i].includes(':')) {
            const [key, ...value] = lines[i].split(': ');
            jsonObj[key.trim()] = value.join(': ');
        } else if (lines[i].trim() === '') { // Once we hit an empty line, the rest is the message content
            jsonObj['Content'] = lines.slice(i+1).join('\n').trim();
            break;
        }
    }
    result.push(jsonObj);
  }

  return result;
}

// .eml can contain multiple boundaries (nested), so we need to find them all
function findEmlBoundaries(text) {
  const boundaryRegex = /boundary=(?:"([^"]+)"|([^;\n\r]+))/gi;
  const boundaryMatches = text.match(boundaryRegex);
  if (boundaryMatches) {
    const boundaries = boundaryMatches.map(match => {
      const boundary = match.replace(/boundary=/i, '');
      return boundary.startsWith('"') && boundary.endsWith('"')
        ? "--" + boundary.slice(1, -1)
        : "--" + boundary;
    });
    return boundaries;
  } else {
    console.log('Boundary not found');
    return [];
  }
}

function stripHtmlTags(html) {
  const plainText = html.replace(/<[^>]+>/g, '');
  return decodeEntities(plainText);
}

function decodeEntities(text) {
  const entities = [
    ['amp', '&'],
    ['apos', '\''],
    ['lt', '<'],
    ['gt', '>'],
    ['quot', '"']
  ];
  
  for (const entity of entities) {
    const [name, value] = entity;
    const regex = new RegExp(`&${name};`, 'g');
    text = text.replace(regex, value);
  }
  
  return text;
}

function createCard(text) {
  const textParagraph = CardService.newTextParagraph().setText(text);
  const section = CardService.newCardSection().addWidget(textParagraph);
  const card = CardService.newCardBuilder().addSection(section);
  return card;
}