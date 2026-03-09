import escapeXmlCharacters from './escapeXmlCharacters.js'
import removeInvalidXmlCharacters from './removeInvalidXmlCharacters.js'

/**
 * Escapes invalid characters in an XML element's text content.
 * @param {string} textContent
 * @returns {string}
 */
export default function sanitizeTextContent(textContent) {
	if (typeof textContent !== 'string') {
		throw new TypeError('Argument must be a string')
	}
	return escapeXmlCharacters(removeInvalidXmlCharacters(textContent), { isAttributeValue: false })
}