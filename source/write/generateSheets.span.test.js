import generateSheets from './generateSheets.js'

describe('generateSheets()', () => {
  it('should detect `span`/`rowSpan` overlap', function() {
    expect(
      () => generateSheets({
        data: [
          [
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
          ]
        ],
        sheetNames: ['Sheet 1'],
        features: [],
        multipleSheetsParameters: true
      })
    ).to.throw('span')
  })
})
