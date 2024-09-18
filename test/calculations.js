const { expect } = require('chai');

describe('Calculations', () => {
    it('1', () => {
        let test = jspreadsheet(root, {
            worksheets: [{
                data: [
                    ['1', '2', '3'],
                    ['4', '5', '6'],
                    ['7', '8', '9'],
                    ['=SUM(B1:B3)', '=SUM(B1:B3)', '=SUM(B1:B3)'],
                ],
                worksheetName: 'sheet1',
            }],
        });

        test[0].insertColumn(1, 1, true);

        expect(test[0].getValue('A4')).to.equal('=SUM(C1:C3)');
        expect(test[0].getValue('B4')).to.equal('');
        expect(test[0].getValue('C4')).to.equal('=SUM(C1:C3)');
        expect(test[0].getValue('D4')).to.equal('=SUM(C1:C3)');
    })
});