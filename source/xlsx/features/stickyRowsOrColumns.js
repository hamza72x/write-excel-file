import findElement from '../../xml/findElement.js'
import getAttributesString from '../../xml/getAttributesString.js'
import getCellCoordinate from '../helpers/getCellCoordinate.js'

export default {
	files: {
		transform: {
			'xl/worksheets/sheet{id}.xml': {
				transform(xml, { stickyRowsCount, stickyColumnsCount, sheetIndex, sheetId }) {
					if (stickyRowsCount || stickyColumnsCount) {
						const paneAttributes = {
							// The vertical position of the split, in 1/20th of a point (twips),
							// or, when frozen, the number of visible rows in the top pane(s).
							ySplit: stickyRowsCount || 0,
							// The horizontal position of the split, in 1/20th of a point (twips),
							// or, when frozen, the number of visible columns in the left pane(s).
							xSplit: stickyColumnsCount || 0,
							// The visible cell in the top-left corner of the bottom-right pane.
							topLeftCell: getCellCoordinate(stickyRowsCount || 0, stickyColumnsCount || 0),
							// `activePane` defines which pane is active in a "split" or "frozen" `state`.
							// Possible values: "bottomLeft", "bottomRight", "topLeft", "topRight".
							activePane: 'bottomRight',
							// `state: "frozen"` indicates that the panes are frozen, meaning the split area
							// (usually top rows or leftmost columns) remains fixed while the rest of the sheet scrolls.
							// Other possible values:
							// "split" — Indicates that the window is split into separate panes that can be scrolled independently, but are not locked in place.
							// "frozenSplit" — A combination used in specific scenarios (often by Excel itself) to manage complex freezing.
							state: 'frozen'
						}

						const sheetViewElement = findElement(xml, 'sheetView')
						if (sheetViewElement) {
							// Add a `<pane/>` element inside the `<sheetView/>` element.
							// If a `<pane/>` element already exists, it will overwrite it.
							const sheetViewInnerXml = xml.slice(sheetViewElement.openingTagEndIndex + 1, sheetViewElement.closingTagStartIndex)
							const paneElement = findElement(sheetViewInnerXml, 'pane')
							const paneXml = `<pane${getAttributesString(paneAttributes)}/>`
							if (paneElement) {
								// Overwrite the existing `<pane/>` element.
								// It doesn't seem to be any sense in patching the existing `<pane/>` element
								// because all of its properties would be rewritten anyway in the process.
								xml =
									xml.slice(0, sheetViewElement.openingTagEndIndex + 1 + paneElement.openingTagStartIndex) +
									paneXml +
									xml.slice(sheetViewElement.openingTagEndIndex + 1 + (paneElement.selfClosingTag ? paneElement.openingTagEndIndex : paneElement.closingTagEndIndex) + 1)
							} else {
								// Add a `<pane/>` element.
								xml = xml.replace('</sheetView>', paneXml + '</sheetView>')
							}
						} else {
							// Add a `<sheetViews/>` element.
							// It isn't supposed to exist when no `<sheetView/>` element is present.
							const sheetViewsElement = findElement(xml, 'sheetViews')
							if (sheetViewsElement) {
								throw new Error(`xl/worksheets/sheet${sheetId}.xml: <sheetViews/> element exists but it doesn't contain any <sheetView/> elements`)
							}

							const sheetViewAttributes = {
								tabSelected: sheetIndex === 0 ? 1 : 0, // The first sheet is selected by default.
								workbookViewId: 0
							}
							const paneXml = `<pane${getAttributesString(paneAttributes)}/>`
							const sheetViewXml = `<sheetView${getAttributesString(sheetViewAttributes)}>${paneXml}</sheetView>`
							const sheetViewsXml = `<sheetViews>${sheetViewXml}</sheetViews>`
							const worksheetElement = findElement(xml, 'worksheet')

							// Add a `<sheetViews/>` element with a child '<sheetView/>` element with a child `<pane/>` element.
							// For some reason, Excel 2007 demands `<sheetViews/>` element to be the first one in a `<worksheet/>` element.
							xml =
								xml.slice(0, worksheetElement.openingTagEndIndex + 1) +
								// For some weird reason, Excel 2007 demands `<sheetViews/>` element to be the first one.
								sheetViewsXml +
								xml.slice(worksheetElement.openingTagEndIndex + 1)
						}
					}

					return xml
				},

				// These parameters will be passed through to the function above.
				parameters: (availableParameters) => {
					const { stickyRowsCount, stickyColumnsCount } = availableParameters
					return { stickyRowsCount, stickyColumnsCount }
				}
			},

			'xl/workbook.xml': {
				transform: (xml, { stickyRowsCount, stickyColumnsCount }) => {
					if (stickyRowsCount || stickyColumnsCount) {
						// Find a `<workbookView/>` element.
						// It is required to exist because it is referenced as `workbookViewId="0"`
						// attribute in `<sheetViews/>` element in `xl/worksheets/sheet{id}.xml` file.
						const workbookViewElement = findElement(xml, 'workbookView')
						if (!workbookViewElement) {
							// `<workbookView/>` element doesn't exist, so it should be created.
							// The parent element for it should be `<bookViews/>`.
							const bookViewsElement = findElement(xml, 'bookViews')
							if (bookViewsElement) {
								throw new Error(`xl/workbook.xml: <bookViews/> element exists but it doesn't contain any <workbookView/> elements`)
							}
							// Add a `<bookViews/>` element with a child '<workbookView/>` element.
							// For some reason, Excel 2007 demands `<bookViews/>` element to go before `<sheets/>` element.
							// And `<sheets/>` element always exists.
							xml = xml.replace('<sheets>', '<bookViews><workbookView/></bookViews>' + '<sheets>')
						}
					}
					return xml
				},

				// These parameters will be passed through to the function above.
				parameters: (availableParameters) => {
					const { stickyRowsCount, stickyColumnsCount } = availableParameters
					return { stickyRowsCount, stickyColumnsCount }
				}
			}
		}
	}
}