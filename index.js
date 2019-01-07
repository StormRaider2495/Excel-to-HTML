"use strict";
var fs = require('fs');
var process = require('process');
var workingDirectory = process.cwd().slice(2);
var XLSX = require('xlsx');
var workbook = XLSX.readFile(process.argv[2]);
var sheets = workbook.Sheets;
var htmlFile = '';
var rowNumber;
var htmlArray;

// Check to make sure user provides argument for command line
if (typeof process.argv[2] === 'undefined') {
	console.log('\n' + 'Error:' + '\n' + 'You must enter the excel file you wish to build tables from as an argument' + '\n' + 'i.e., node toTable.js resolutions.xlsx');
	return;
} else {
	// Check that the file is the correct type
	if (process.argv[2].slice(-4) !== 'xlsx') {
		console.log('\n' + 'This program will only convert xlsx files' + '\n' + 'Please enter correct file type');
		return;
	} else {
		// Create the HTML file name to write the table to
		var fileName = process.argv[2];
		var newFileName = fileName.slice(0, -4) + 'html';
	}
}

// start HTML file
htmlFile = '<!DOCTYPE html>' + '\n' +
'<html lang="en">' + '\n' +
'<head>' + '\n' +
    '<meta charset="UTF-8">' + '\n' +
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">' + '\n' +
    '<meta http-equiv="X-UA-Compatible" content="ie=edge">' + '\n' +
    '<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/css/bootstrap.min.css" integrity="sha384-MCw98/SFnGE8fJT3GXwEOngsV7Zt27NXFoaoApmYm81iuXoPkFOJwJ8ERdknLPMO" crossorigin="anonymous">' + '\n' +
	'<title>' + fileName.slice(0, -5) + '</title>' + '\n' +
	'<style>' +
	'.table-view { height: 450px; overflow: auto; }' +
	'</style>' +
'</head>' + '\n' +
'<body>' + '\n';

function getPosition(string, subString, index) {
   return string.split(subString, index).join(subString).length;
}
// Iterate through each worksheet in the workbook
var sheetNumber = 0;
for (var sheet in sheets) {
	sheetNumber++;
	htmlFile += '<h2> Sheet ' + sheetNumber+ ':' + sheet + '</h2>';
	// Start building a new table if the worksheet has entries
	if (typeof sheet !== 'undefined') {
		htmlFile += ' <div class="container-fluid table-view"> <table summary="" class="table">' + '\n' + '<thead class="thead-dark">';		
		// Iterate over each cell value on the sheet

		/*
		 * TODO: the module omits blank cells
		 * To add them check between last cell and current cell to check if cells are missing.
		 * HINT: in cells the numbers is always the same for a row, alphabets represent columns, which are skipped.
		*/
		var lastCell = '';
		for (var cell in sheets[sheet]) {	
			// catch first value as lastCell		
			lastCell = lastCell === '' ? cell : lastCell;

			// Protect against undefined values			
			if (typeof sheets[sheet][cell].w !== 'undefined') {
				//The first row in the table
				if (cell === 'A1') {
					htmlFile += '\n' + '<tr>' + '\n' + '<th>' + sheets[sheet][cell].w.replace('& ', '&amp;').replace('-', '&ndash;').replace('–', '&mdash;') + '</th>';
				} else {
					//The second row in the table closes the thead element
					if (cell === 'A2') {
						htmlFile += '\n' + '</tr>' + '\n' + '</thead>' + '\n' + '<tr>' + '\n' + '<th>' + sheets[sheet][cell].w.replace('& ', '&amp;').replace('-', '&ndash;').replace('–', '&mdash;') + '</th>';
					} else {
						// The first cell in each row
						if (cell.slice(0, 1) === 'A') {
							htmlFile += '\n' + '</tr>' + '\n' + '<tr>' + '\n' + '<th>' + sheets[sheet][cell].w.replace('& ', '&amp;').replace('-', '&ndash;').replace('–', '&mdash;') + '</th>';
							//All the other cells
						} else {
							htmlFile += '\n' + '<td>' + sheets[sheet][cell].w.replace('& ', '&amp;').replace('-', '&ndash;').replace('–', '&mdash') + '</td>';
						}
					}
				}
			}
		}
		// Close the table
		htmlFile += '\n' + '</tr>' + '\n' + '</table>' +'\n' + '</div>' + '\n';
	}
	/*console.log(sheets[sheet]['!merges']);
	sheets[sheet]['!merges'].forEach(function(merge, index) {
		//console.log(merge);
		rowNumber = (getPosition(htmlFile, '<th>', (merge.s.r+1)) + 3);
		console.log(rowNumber);
		htmlArray = htmlFile.split('');
		htmlArray.splice(rowNumber, 0, ' colspan="3"');
		htmlFile = htmlArray.join('');
	});*/
}
// Close the file
htmlFile += '</body>' + '\n' + '</html>';

// Write htmlFile variable to the disk with newFileName as the name
fs.writeFile(newFileName, htmlFile, (err) => {
	if (err) throw err;
	console.log('\n' +'Your tables have been created in', newFileName);
});