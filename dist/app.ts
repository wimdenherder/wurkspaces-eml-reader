
import { FileSelectedEvent } from './types/google-drive-addon';

function onHomepage() {
  return createCard('Welcome to EML Reader.\nPlease select an .EML file');
}

// eslint-disable-next-line
function onFileSelected(e: FileSelectedEvent) {
  const MIMETYPE_EML = 'message/rfc822';
  const item = e.drive.selectedItems[0]; // take first item (when multiple files are selected)

  if(isBiggerThan4GB(item.id))
    return createCard("Error message: file is bigger than 4GB.");
  if(item.mimeType !== MIMETYPE_EML)
    return onHomepage();
  
  const emlContent = getFileContent(item.id);
  const description = describeEml(emlContent);
  return createCard(description);
}

function isBiggerThan4GB(fileId: string) {
  const maxSize = 4 * 1024 * 1024 * 1024;
  const file = DriveApp.getFileById(fileId);
  const fileSize = file.getSize();
  return fileSize > maxSize;
}

function getFileContent(id: string) {
  try {
    const file = DriveApp.getFileById(id);
    return file.getBlob().getDataAsString();
  } catch (error) {
    throw new Error("Failed to get file contents: " + error.message);
  }
}

function describeEml(emlContent: string): string {
  const emailData = convertEMLToJSON(emlContent);

  const isAttachment = 
    part => part['Content-Disposition']?.startsWith("attachment");

  const isPlainTextEmail =
    part => !isAttachment(part) && part['Content-Type']?.startsWith('text/plain');

  const isHtmlEmail =
    part => !isAttachment(part) && part['Content-Type']?.startsWith('text/html');

  const emailPlainText = emailData.find(isPlainTextEmail);
  const emailHtml = emailData.find(isHtmlEmail);
  
  const content = emailPlainText?.Content || stripHtmlTags(emailHtml?.Content || '');
  
  const attachments = emailData.filter(isAttachment);

  return `Subject: ${emailData[0].Subject}\nFrom: ${emailData[0].From}\nDate: ${emailData[0].Date}\nAttachments: ${attachments.length}\n\nBody: \n\n${content}`;
}

function convertEMLToJSON(text: string) {
  const result = [];
  const boundaries = findEmlBoundaries(text);
  const boundariesSplitter = new RegExp(boundaries.map(x => x + "\r?\n?").join('|'));
  const notEmpty = x => x;
  const parts = boundaries.length > 0
                  ? text.split(boundariesSplitter).filter(notEmpty)
                  : [text];
  
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
function findEmlBoundaries(text: string) {
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

function stripHtmlTags(html: string) {
  const plainText = html.replace(/<[^>]+>/g, '');
  return decodeEntities(plainText);
}

function decodeEntities(text: string) {
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

function createCard(text: string) {
  const textParagraph = CardService.newTextParagraph().setText(text);
  const section = CardService.newCardSection().addWidget(textParagraph);
  const card = CardService.newCardBuilder().addSection(section);
  return [card.build()];
}