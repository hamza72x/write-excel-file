import escapeXmlCharacters from './escapeXmlCharacters.js'
import removeInvalidXmlCharacters from './removeInvalidXmlCharacters.js'

/**
 * Escapes invalid characters in an XML attribute value.
 * @param {string} attributeValue
 * @returns {string}
 */
export default function sanitizeAttributeValue(attributeValue) {
	if (typeof attributeValue !== 'string') {
		throw new TypeError('Argument must be a string')
	}
	return escapeXmlCharacters(removeInvalidXmlCharacters(attributeValue), { isAttributeValue: true })
}