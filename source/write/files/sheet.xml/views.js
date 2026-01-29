import getCellCoordinate from '../../helpers/getCellCoordinate.js'
import getAttributesString from '../../../xml/getAttributesString.js'

export default function generateViews({
	stickyRowsCount,
	stickyColumnsCount,
	showGridLines,
	rightToLeft,
	sheetIndex
}) {
	// Custom "plug-ins" might modify the `<views/>` element contents.
	// For example, there might be a plug-in that inserts a `zoomScale` attribute.
	// Because of that, `<views/>` element is always created.
	//
	// if (
	// 	!stickyRowsCount &&
	// 	!stickyColumnsCount &&
	// 	!(showGridLines === false) &&
	// 	// typeof zoomScale !== 'number' &&
	// 	!rightToLeft
	// ) {
	// 	return ''
	// }

	let views = ''

	const sheetViewAttributes = {
		tabSelected: sheetIndex === 0 ? 1 : 0, // The first sheet is selected by default.
		workbookViewId: 0
	}

	if (showGridLines === false) {
		sheetViewAttributes.showGridLines = false
	}

	// if (zoomScale) {
	// 	sheetViewAttributes.zoomScale = zoomScale
	// }

	if (rightToLeft) {
		sheetViewAttributes.rightToLeft = 1
	}

	const paneAttributes = {
		ySplit: stickyRowsCount || 0,
		xSplit: stickyColumnsCount || 0,
		topLeftCell: getCellCoordinate(stickyRowsCount || 0, stickyColumnsCount || 0),
		// `activePane` defines which pane is currently active in a split or frozen worksheet view.
		activePane: 'bottomRight',
		// `state: "frozen"` indicates that the panes are frozen, meaning the split area
		// (usually top rows or leftmost columns) remains fixed while the rest of the sheet scrolls.
		// Other possible values:
		// "split" — Indicates that the window is split into separate panes that can be scrolled independently, but are not locked in place.
		// "frozenSplit" — A combination used in specific scenarios (often by Excel itself) to manage complex freezing.
		state: 'frozen'
	}

	views += '<sheetViews>'
	views += `<sheetView${getAttributesString(sheetViewAttributes)}>`
	views += `<pane${getAttributesString(paneAttributes)}/>`
	views += '</sheetView>'
	views += '</sheetViews>'

	return views
}