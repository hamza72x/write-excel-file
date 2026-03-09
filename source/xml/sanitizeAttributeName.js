import escapeXmlCharacters from './escapeXmlCharacters.js'
import removeInvalidXmlCharacters from './removeInvalidXmlCharacters.js'

/**
 * Escapes invalid characters in an XML attribute name.
 * @param {string} attributeName
 * @returns {string}
 */
export default function sanitizeAttributeName(attributeName) {
	if (typeof attributeName !== 'string') {
		throw new TypeError('Argument must be a string')
	}
	return escapeXmlCharacters(removeInvalidXmlCharacters(attributeName), { isAttributeValue: false })
}