const extractor = require('pst-extractor');
const { PSTFile, PSTFolder } = extractor;
const testFolder = './mails_list/';
const fs = require('fs');
let depth = -1;
let col = 0;

// Require library
var xl = require('excel4node');

// Create a new instance of a Workbook class
var wb = new xl.Workbook();

// Add Worksheets to the workbook
var ws = wb.addWorksheet('Mails');

// Create a reusable style
var style = wb.createStyle({
  font: {
    color: '#FF0800',
    size: 14
  },
  numberFormat: '$#,##0.00; ($#,##0.00); -'
});

// Set value of cell A1 to 100 as a number type styled with paramaters of style
ws.cell(1, 1)
  .string('Emails')
  .style(style);

let mailsList = [];

// Get all pst files from 'mails_list' folder, and process them
fs.readdirSync(testFolder).forEach(file => {
  const pstFile = new PSTFile('./mails_list/' + file);
  processFolder(pstFile.getRootFolder());
});

// After we proccess all mails, remove duplicateds

uniqueMails = removeDuplicatedFromArray(mailsList);
uniqueMails.map((mailItem, i) => {
  ws.cell(i + 2, 1).string(mailItem);
  if (uniqueMails.length == i + 1) {
    wb.write('list_mails.xlsx');
  }
});

/**
 * Walk the folder tree recursively and process emails.
 * @param {PSTFolder} folder
 */
function processFolder(folder) {
  depth++;
  if (folder.hasSubfolders) {
    let childFolders = folder.getSubFolders();
    childFolders.map((childFolder, i) => {
      processFolder(childFolder);
    });
  }

  // and now the emails for this folder
  if (folder.contentCount > 0) {
    depth++;
    let email = folder.getNextChild();
    while (email != null && email.senderName != null) {
      mailsList.push(email.senderEmailAddress);
      email = folder.getNextChild();
    }
    depth--;
  }
  depth--;
}

function removeDuplicatedFromArray(array) {
  let unique = {};
  array.forEach(function(i) {
    if (!unique[i]) {
      unique[i] = true;
    }
  });
  return Object.keys(unique);
}
