import findElements from './findElements.js'

/**
 * Finds a single element in valid XML markup.
 * @param {string} xml — XML markup
 * @param {string} tagName — The name of the element to find
 * @returns {object[]} — A list of found elements, each represented by an object with propeties: `{ openingTagXml: string, openingTagStartIndex: number, openingTagEndIndex: number, openingTagAttributes: object, selfClosingTag?: boolean, closingTagXml?: string, closingTagStartIndex?: number, closingTagEndIndex?: number }`
 */
export default function findElement(xml, tagName) {
	return findElements(xml, tagName)[0]
}