import $attr from '../../xml/sanitizeAttributeValue.js'
import getXlsxColorForHexColor from './getXlsxColorForHexColor.js'

export function getBorderXml(border) {
	const {
		left,
		right,
		top,
		bottom
	} = border
	let xml = ''
	xml += '<border>'
	xml += getSideBorderXml('left', left)
	xml += getSideBorderXml('right', right)
	xml += getSideBorderXml('top', top)
	xml += getSideBorderXml('bottom', bottom)
	xml += '<diagonal/>'
	xml += '</border>'
	return xml
}

function getSideBorderXml(side, { style, color }) {
	if (color && !style) {
		style = 'thin'
	}
	const hasChildren = Boolean(color)
	return `<${side}` +
		(style ? ` style="${$attr(style)}"` : '') +
		(hasChildren ? '>' : '/>') +
		(color ? `<color rgb="${$attr(getXlsxColorForHexColor(color))}"/>` : '') +
		(hasChildren ? `</${side}>` : '')
}