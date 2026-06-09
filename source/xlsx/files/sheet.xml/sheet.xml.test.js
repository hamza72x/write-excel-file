import { describe, it } from 'mocha'
import { expect } from 'chai'

import generateSheetXml from './sheet.xml.js'

describe('sheet.xml', () => {
  it('should generate `sheet.xml` markup', () => {
  	const colsXml = ''

  	const row1Xml = '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1000</v></c></row>'
  	const row2Xml = '<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>2000</v></c></row>'
    const rowsXml = row1Xml + row2Xml

    // const paneXml = '<pane ySplit="0" xSplit="0" topLeftCell="A1" activePane="bottomRight" state="frozen"/>'
    // const sheetViewsXml = '<sheetViews><sheetView tabSelected="0" workbookViewId="0">' + paneXml + '</sheetView></sheetViews>'
    // const sheetViewsXml = '<sheetViews><sheetView tabSelected="0" workbookViewId="0"/></sheetViews>'
    const sheetViewsXml = ''

    const expectedXmlBeforeRowsXml = '<?xml version="1.0" ?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
        sheetViewsXml +
        colsXml +
        '<sheetData>'

    const expectedXmlAfterRowsXml =
        '</sheetData>' +
        '<drawing r:id="rId-drawing-1"/>' +
      '</worksheet>'

    const expectedXml = expectedXmlBeforeRowsXml + rowsXml + expectedXmlAfterRowsXml

    const sheetData = [
      [
        {
          value: 'Test',
          type: String,
          // width: 150
        },
        {
          value: 1000,
          type: Number,
          // width: 60
        }
      ],
      [
        {
          value: 'Test 2',
          type: String,
          // width: 150
        },
        {
          value: 2000,
          type: Number,
          // width: 60
        }
      ]
    ]

    expect(
      generateSheetXml(
        {
          findOrCreateSharedString(string) {
            if (string === 'Test') {
              return 0
            }
            if (string === 'Test 2') {
              return 1
            }
            throw new Error(`Unknown shared string: ${string}`)
          },
          sheetData,
          sheetIndex: 0,
          sheetId: 1,
          sheetOptions: {}
        },
        { streaming: false },
        []
      )
    ).to.equal(expectedXml)
  })

  it('should generate `sheet.xml` markup (`streaming: true`)', () => {
  	const colsXml = ''

  	const row1Xml = '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1000</v></c></row>'
  	const row2Xml = '<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>2000</v></c></row>'
    const rowsXml = row1Xml + row2Xml

    // const paneXml = '<pane ySplit="0" xSplit="0" topLeftCell="A1" activePane="bottomRight" state="frozen"/>'
    // const sheetViewsXml = '<sheetViews><sheetView tabSelected="0" workbookViewId="0">' + paneXml + '</sheetView></sheetViews>'
    // const sheetViewsXml =
    //  '<sheetViews><sheetView tabSelected="0" workbookViewId="0"/></sheetViews>'
    const sheetViewsXml = ''

    const expectedXmlBeforeRowsXml = '<?xml version="1.0" ?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
        sheetViewsXml +
        colsXml +
        '<sheetData>'

    const expectedXmlAfterRowsXml =
        '</sheetData>' +
        '<drawing r:id="rId-drawing-1"/>' +
      '</worksheet>'

    const expectedXml = expectedXmlBeforeRowsXml + rowsXml + expectedXmlAfterRowsXml

    const sheetData = [
      [
        {
          value: 'Test',
          type: String,
          // width: 150
        },
        {
          value: 1000,
          type: Number,
          // width: 60
        }
      ],
      [
        {
          value: 'Test 2',
          type: String,
          // width: 150
        },
        {
          value: 2000,
          type: Number,
          // width: 60
        }
      ]
    ]

    const {
      xmlBeforeRowsXml,
      xmlAfterRowsXml,
      rowCount,
      getRowXml
    } = generateSheetXml(
      {
        findOrCreateSharedString(string) {
          if (string === 'Test') {
            return 0
          }
          if (string === 'Test 2') {
            return 1
          }
          throw new Error(`Unknown shared string: ${string}`)
        },
        sheetData,
        sheetIndex: 0,
        sheetId: 1,
        sheetOptions: {}
      },
      { streaming: true },
      []
    )

    expect(xmlBeforeRowsXml).to.equal(expectedXmlBeforeRowsXml)
    expect(xmlAfterRowsXml).to.equal(expectedXmlAfterRowsXml)
    expect(rowCount).to.equal(2)
    expect(() => getRowXml(-1)).to.throw('out of bounds')
    expect(() => getRowXml(2)).to.throw('out of bounds')
    expect(getRowXml(0)).to.equal(row1Xml)
    expect(getRowXml(1)).to.equal(row2Xml)
  })

  it('should detect `columnSpan`/`rowSpan` overlap', () => {
    expect(() => {
      generateSheetXml(
        {
          findOrCreateSharedString(string) {
            return 0
          },
          sheetData: [
            [
              {
                value: 'Text',
                columnSpan: 3
              },
              null,
              {
                value: 'Text'
              }
            ]
          ],
          sheetIndex: 0,
          sheetId: 1,
          sheetOptions: {}
        },
        { streaming: false },
        []
      )
    }).to.throw('columnSpan')
  })

  it('should detect `span`(legacy property)/`rowSpan` overlap', () => {
    expect(() => {
      generateSheetXml(
        {
          findOrCreateSharedString(string) {
            return 0
          },
          sheetData: [
            [
              {
                value: 'Text',
                span: 3
              },
              null,
              {
                value: 'Text'
              }
            ]
          ],
          sheetIndex: 0,
          sheetId: 1,
          sheetOptions: {}
        },
        { streaming: false },
        []
      )
    }).to.throw('columnSpan')
  })
})
