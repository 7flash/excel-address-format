var assert = require('assert');
var address = require('./index');

describe("ExcelAddressFormat", function() {
	it('toAlphabetFormat valid input', function() {
		assert.equal(address.toAlphabetFormat('R4C1'), 'A4');
		assert.equal(address.toAlphabetFormat('R1C1'), 'A1');
	});

	it('toAlphabetFormat invalid input', function() {
		assert.equal(address.toAlphabetFormat('B5'), false);
	});

	it('toNumberFormat valid input', function() {
		assert.equal(address.toNumberFormat('A4'), 'R4C1');
		assert.equal(address.toNumberFormat('B5'), 'R5C2');
	});

	it('toNumberFormat invalid input', function() {
		assert.equal(address.toNumberFormat('R4C1'), false);
	});

	it('isNumberAddress should check if address is numeric', function() {
		assert.equal(address.isNumberAddress('R4C1'), true);
		assert.equal(address.isNumberAddress('R100C200'), true);
		assert.equal(address.isNumberAddress('A1'), false);
	});

	it('isAlphabetAddress should check if address is alphabetic', function() {
		assert.equal(address.isAlphabetAddress('A5'), true);
		assert.equal(address.isAlphabetAddress('AA100'), true);
		assert.equal(address.isAlphabetAddress('R4C1'), false);
	});
	
	it('letterToNumber should convert excel letter to number representation', function() {
		assert.equal(address.letterToNumber('A'), 1);
		assert.equal(address.letterToNumber('AA'), 27);
		assert.equal(address.letterToNumber('Z'), 26);
		assert.equal(address.letterToNumber('AZ'), 52);
		assert.equal(address.letterToNumber('100'), -1);
	});

	it('numberToLetter should convert number to excel letter', function() {
		assert.equal(address.numberToLetter(0), -1);
		assert.equal(address.numberToLetter(1), 'A');
		assert.equal(address.numberToLetter(27), 'AA');
		assert.equal(address.numberToLetter('A'), -1);
	});

	it('getRow should return row from cell', function() {
		assert.equal(address.getRow('0'), -1);
		assert.equal(address.getRow('R4C1'), 4);
		assert.equal(address.getRow('A5'), 5);
	});

	it('getColumn should return column from cell', function() {
		assert.equal(address.getColumn('0'), -1);
		assert.equal(address.getColumn('B5'), 'B');
		assert.equal(address.getColumn('R4C1'), '1');
		assert.equal(address.getColumn('AA100'), 'AA');
	});
});
