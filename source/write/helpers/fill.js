import $attr from '../../xml/sanitizeAttributeValue.js'
import getXlsxColorForHexColor from './getXlsxColorForHexColor.js'

export function hasFill(fill) {
	const {
		backgroundColor,
		fillPatternStyle,
		fillPatternColor
	} = fill
	return backgroundColor || (fillPatternStyle && fillPatternColor)
}

export function getFillXml(fill) {
	const {
		backgroundColor,
		fillPatternStyle,
		fillPatternColor
	} = fill

	const isSolidFill = !fillPatternStyle || fillPatternStyle === 'solid'

	if (!hasFill(fill)) {
		return '<fill>' + '<patternFill patternType="none"/>' + '</fill>'
	}

	let xml = '<fill>'
	// A non-pattern solid background is still applied via a "solid" "fill pattern" for some reason.
	// It's not clear why do they do it in such a way.
	// Perhaps it's an established practice in Excel world or something.
	xml += `<patternFill patternType="${isSolidFill ? 'solid' : fillPatternStyle}">`
	// If it's a "fill pattern" then set its foreground color.
	// If it's not a "fill pattern" and just a solid background then set the foreground color
	// of the fill same as the background color.
	xml += `<fgColor rgb="${$attr(getXlsxColorForHexColor(isSolidFill ? backgroundColor : fillPatternColor))}"/>`
	// `<bgColor indexed="64"/>` means "black background".
	// It doesn't really matter if it's all black when a "solid" "fill pattern" exists above it anyway.
	// It's not clear why do they insert such redundant `<bgColor/>` in such cases.
	// Perhaps it's an established practice in Excel world or something.
	xml += `<bgColor ${isSolidFill ? 'indexed="64"' : 'rgb="' + $attr(getXlsxColorForHexColor(backgroundColor)) + '"'}/>`
	xml += '</patternFill>'
	// Close the `<fill>` element.
	xml += '</fill>'
	// Return the XML.
	return xml
}