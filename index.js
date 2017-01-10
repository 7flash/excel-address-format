var XLSX = require('xlsx-styles');
var workbook = XLSX.readFile('input/test.XLSX');
var sheet_name_list = workbook.SheetNames;

var alphabetAddressPattern = /^([A-Z]+)(\d)$/;
var numberAddressPattern = /^R((\d)+)C((\d)+)$/;

function toAlphabetFormat(numberAddress) {
	if(!isNumberAddress(numberAddress)) return false;

	var row = parseInt(getRow(numberAddress));

	var column = parseInt(getColumn(numberAddress)) - 1;
	var columnLetter = numberToLetter(column);

	var alphabetFormattedIndex = columnLetter.concat(row);
	return alphabetFormattedIndex;
}

function toNumberFormat(alphabetAddress) {
	if(!isAlphabetAddress(alphabetAddress)) return false;

	var row = parseInt(getRow(alphabetAddress));

	var columnLetter = getColumn(alphabetAddress);
	var column = parseInt(letterToNumber(columnLetter)) + 1;

	var numberFormattedIndex = 'R'+row+'C'+column;
	return numberFormattedIndex;
}

function letterToNumber(letter) {
	var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';	
	var sum = 0;
	for(var i = 0; i < alphabet.length; i++) {
		sum *= 26;
		sum += (alphabet.charCodeAt(i) - ("A".charCodeAt(0)-1));
	}
	return sum;
}

function numberToLetter(number) {
    var dividend = number;
    var columnName = "";
    var modulo;

    while (dividend > 0)
    {
        modulo = (dividend - 1) % 26;
        columnName = String.fromCharCode(65 + modulo).toString() + columnName;
        dividend = parseInt((dividend - modulo) / 26);
    } 
    return columnName;
}

function isNumberAddress(address) {
	return numberAddressPattern.test(address) === true;
}

function isAlphabetAddress(address) {
	return alphabetAddressPattern.test(address) === true;
}

function getColumn(address) {
	if(isAlphabetAddress(address)) {
		var match = address.match(alphabetAddressPattern);
		return match[1];
	} else if(isNumberAddress(address)) {
		var match = address.match(numberAddressPattern);
		return match[2];
	}
	return -1;
}

function getRow(address) {
	if(isAlphabetAddress(address)) {
		var match = address.match(alphabetAddressPattern);
		return match[2];
	} else if(isNumberAddress(address)) {
		var match = address.match(numberAddressPattern);
		return match[1];
	}
	return -1;
}

function getBlocks(worksheet) {
	console.log(toNumberFormat("B5"));
	console.log(toAlphabetFormat("R4C1"));
	return [];

	var bounds = worksheet['!ref'];
	var numberFormattedBounds = toNumberFormat(bounds);
	var maxColumn = getColumn(numberFormattedBounds);
	var maxRow = getRow(numberFormattedBounds);

	var blocks = {};
	var startColumn = 5;
	var startRow = 8;
	var columnSize = 5; 
	var rowSize = 7;

	var currentColumn = startColumn;
	var currentRow = startRow;
	while(true) {
		var block = {};
		
		currentRow += rowSize;
	}
}

function parseBlock() {

}

sheet_name_list.forEach(function(y) {
	var worksheet = workbook.Sheets[y];
	getBlocks(worksheet, function(blocks) {
		blocks.forEach(item, function(block) {
			
		});
	});
	/*
	var index = cellIndexFormat('R1C1') + ':' + cellIndexFormat('R1C2');
	console.log(index);
	console.log(worksheet[index].v);
	*/
	console.log(worksheet);
});
