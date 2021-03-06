exports.toLetterFormat = toLetterFormat;
exports.toNumberFormat = toNumberFormat;
exports.isNumberAddress = isNumberAddress;
exports.isAlphabetAddress = isAlphabetAddress;
exports.letterToNumber = letterToNumber;
exports.numberToLetter = numberToLetter;
exports.getRow = getRow;
exports.getColumn = getColumn;
exports.generateLetterAddress = generateLetterAddress;
exports.generateNumberAddress = generateNumberAddress;

var alphabetAddressPattern = /^([A-Z]+)(\d+)$/;
var numberAddressPattern = /^R(\d+)C(\d+)$/;

function generateLetterAddress(row, column) {
	if(!Number.isInteger(row) || !Number.isInteger(column)) return false;
	if(row < 0 || column < 0) return false;
	return numberToLetter(column).concat(row);
}

function generateNumberAddress(row, column) {
	if(!Number.isInteger(row) || !Number.isInteger(column)) return false;
	if(row < 0 || column < 0) return false;
	return "R"+row+"C"+column;
}

function toLetterFormat (address) {
	if(!isNumberAddress(address)) {
		if(isAlphabetAddress(address)) {
			return address;
		} else {
			return false;
		}
	}
	
	var row = parseInt(getRow(address));
	
	var column = parseInt(getColumn(address));
	var columnLetter = numberToLetter(column);

	var alphabetFormattedIndex = columnLetter.concat(row);
	return alphabetFormattedIndex;
}

function toNumberFormat (address) {
	if(!isAlphabetAddress(address)) {
		if(isNumberAddress(address)) {
			return address;
		} else {
			return false;
		}
	}

	var row = parseInt(getRow(address));
	
	var columnLetter = getColumn(address);
	var column = parseInt(letterToNumber(columnLetter));

	var numberFormattedIndex = 'R'+row+'C'+column;
	return numberFormattedIndex;
}

function letterToNumber(letter) {
	if(!/^[A-Z]+$/.test(letter)) return -1;
	var sum = 0;
	for(var i = 0; i < letter.length; i++) {
		sum *= 26;
		sum += (letter.charCodeAt(i) - ("A".charCodeAt(0)-1));
	}
	return sum;
}

function numberToLetter(number) {
    if(!/^\d+$/.test(number) || number == 0) return -1;
    var dividend = number;
    var result = "";
    var modulo;

    while (dividend > 0)
    {
        modulo = (dividend - 1) % 26;
        result = String.fromCharCode(65 + modulo).toString() + result;
        dividend = parseInt((dividend - modulo) / 26);
    } 
    return result;
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
