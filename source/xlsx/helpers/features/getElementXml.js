import getOpeningTagMarkup from '../../../xml/getOpeningTagMarkup.js'
import getClosingTagMarkup from '../../../xml/getClosingTagMarkup.js'
import getSelfClosingTagMarkup from '../../../xml/getSelfClosingTagMarkup.js'

export default function getElementXml(
	fileName,
	tagName,
	attributes, // Could be `undefined` or `null`
	innerXml, // Could be `undefined` or `null`
	index, // Could be `undefined`
	properties,
	sheetOptionsOrSheetsOptions,
	features
) {
	for (const feature of features) {
		const transform = feature.files && feature.files.transform && feature.files.transform[fileName]
		if (transform) {
			if (transform.transformElementAttributes) {
				attributes = transform.transformElementAttributes(
					tagName,
					attributes || NO_ATTRIBUTES,
					index,
					sheetOptionsOrSheetsOptions,
					properties
				)
				// Validate that a developer didn't forget to return a value from the function.
				if (attributes === undefined) {
					throw new Error('`transformElementAttributes()` can\'t return `undefined`')
				}
			}
			if (transform.transformElementContent) {
				innerXml = transform.transformElementContent(
					tagName,
					innerXml || null,
					index,
					sheetOptionsOrSheetsOptions,
					properties
				)
				// Validate that a developer didn't forget to return a value from the function.
				if (innerXml === undefined) {
					throw new Error('`transformElementContent()` can\'t return `undefined`')
				}
			}
		}
	}

	if (innerXml) {
		return getOpeningTagMarkup(tagName, attributes) + innerXml + getClosingTagMarkup(tagName)
	} else {
		return getSelfClosingTagMarkup(tagName, attributes)
	}
}

const NO_ATTRIBUTES = {}