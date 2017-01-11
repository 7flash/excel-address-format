var assert = require('assert');
var address = require('./index');

describe('Excel Address Converter', function() {
	describe('CSF to LBSF conversion', function() {
		it('should return correct converted values from CSF to LBSF', function() {
			assert.equal(address.toLetterFormat('R1C1'), 'A1');
			assert.equal(address.toLetterFormat('R100C100'), 'CV100');
		});

		it('should return not-changed LBSF values', function() {
			assert.equal(address.toLetterFormat('B5'), 'B5');
		});

		it('should return false for wrong values', function() {
			assert.equal(address.toLetterFormat('XXX'), false);
		});
	});

	describe('LBSF to CSF conversion', function() {
		it('should return correct converted values from LBSF to CSF', function() {
			assert.equal(address.toNumberFormat('A4'), 'R4C1');
			assert.equal(address.toNumberFormat('B5'), 'R5C2');
			assert.equal(address.toNumberFormat('GY52'), 'R52C207');
		});

		it('should return not-changed CSF values', function() {
			assert.equal(address.toNumberFormat('R1C1'), 'R1C1');
		});

		it('should return false for wrong values', function() {
			assert.equal(address.toNumberFormat('XXX'), false);
		});
	});

	describe('Checks whether LBSF or CSF input', function() {
		describe('Checks if input conforms to CSF or not', function() {
			it('should return true for CSF input', function() {
				assert.equal(address.isNumberAddress('R4C1'), true);
				assert.equal(address.isNumberAddress('R100C200'), true);
			});

			it('should return false for not-CSF input', function() {
				assert.equal(address.isNumberAddress('A1'), false);
			});
		});

		describe('Checks if input conforms to LBSF or not', function() {
			it('should return true for LBSF input', function() {
				assert.equal(address.isAlphabetAddress('A5'), true);
				assert.equal(address.isAlphabetAddress('AA100'), true);
			});

			it('should return false for not-LBSF input', function() {
				assert.equal(address.isAlphabetAddress('R4C1'), false);
			});
		});
	});

	describe('Excel letter to number conversion', function() {
		it('should convert excel letter to number representation', function() {
			assert.equal(address.letterToNumber('A'), 1);
			assert.equal(address.letterToNumber('AA'), 27);
			assert.equal(address.letterToNumber('Z'), 26);
			assert.equal(address.letterToNumber('AZ'), 52);
		});

		it('should return -1 for not-excel-letters', function() {
			assert.equal(address.letterToNumber('100'), -1);
		});
	});

	describe('Number to excel letter conversion', function() {
		it('should convert number to excel-letter representation', function() {
			assert.equal(address.numberToLetter(1), 'A');
			assert.equal(address.numberToLetter(27), 'AA');
		});

		it('should return -1 for not-number', function() {
			assert.equal(address.numberToLetter(0), -1);
			assert.equal(address.numberToLetter('A'), -1);
		});
	});

	describe('Get parts of address', function() {
		describe('Get row from address', function() {
			it('should return row from CSF address', function() {
				assert.equal(address.getRow('R4C1'), 4);
			});

			it('should return row from LBSF address', function() {
				assert.equal(address.getRow('A5'), 5);
			});

			it('should return -1 for wrong address', function() {
				assert.equal(address.getRow('0'), -1);
			});
		});

		describe('Get column from address', function() {
			it('should return column from CSF address', function() {
				assert.equal(address.getColumn('R4C1'), '1');
			});

			it('should return column from LBSF address', function() {
				assert.equal(address.getColumn('B5'), 'B');
				assert.equal(address.getColumn('AA100'), 'AA');
			});

			it('should return -1 for wrong address', function() {
				assert.equal(address.getColumn('0'), -1);
			});
		});
	});
});
