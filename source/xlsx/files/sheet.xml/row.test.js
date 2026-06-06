import { describe, it } from 'mocha'
import { expect } from 'chai'

import getElementXml from '../../helpers/features/getElementXml.js'

import generateRow from './row.js'

describe('generateRow()', () => {
  it('Should create one row from given data', () => {
    expect(generateRow(
      tag,
      [
        { value: 'Test', type: String },
        { value: 1000, type: Number }
      ],
      0,
      {
        findOrCreateSharedString(string) {
          return 0
        },
        sheetIndex: 0,
        sheetId: 1,
        sheetOptions: {},
        features: []
      }
    )).to.equal(
      // '<row r="1"><c r="A1" t="inlineStr"><is><t>Test</t></is></c><c r="B1"><v>1000</v></c></row>'
    	'<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1000</v></c></row>'
    )
  })
})

function tag(tagName, attributes, innerXml) {
  return getElementXml('xl/worksheets/sheet{id}.xml', tagName, attributes, innerXml, 0, {}, {}, [])
}