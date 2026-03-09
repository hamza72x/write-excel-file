/**
 * Finds all elements in valid XML markup.
 * @param {string} xml — XML markup
 * @param {string} tagName — The name of the element to find
 * @returns {object|undefined} — A found element, represented by an object with propeties: `{ openingTagXml: string, openingTagStartIndex: number, openingTagEndIndex: number, openingTagAttributes: object, selfClosingTag?: boolean, closingTagXml?: string, closingTagStartIndex?: number, closingTagEndIndex?: number }`
 */
export default function findElements(xml, tagName) {
	const openingTagRegExp = new RegExp('<' + tagName + '(?:\\s+[^>]+|/)?>', 'g')
	const closingTagRegExp = new RegExp('</' + tagName + '>', 'g')

	const results = []

	let match
	while ((match = openingTagRegExp.exec(xml)) !== null) {
		const attributes = {}
		const attributeRegExp = new RegExp('\\s+([^\\s=>]+)(?:="([^\\s=>]+)")', 'g')
		let attributeMatch
		while ((attributeMatch = attributeRegExp.exec(match[0])) !== null) {
			attributes[attributeMatch[1]] = attributeMatch[2]
		}
		const result = {
			openingTagXml: match[0],
			openingTagStartIndex: match.index,
			openingTagEndIndex: match.index + match[0].length - 1,
			openingTagAttributes: attributes,
			selfClosingTag: undefined,
			closingTagXml: undefined,
			closingTagStartIndex: undefined,
			closingTagEndIndex: undefined
		}
		if (result.openingTagXml[result.openingTagXml.length - 2] === '/') {
			result.selfClosingTag = true
			results.push(result)
		} else {
			closingTagRegExp.lastIndex = result.openingTagEndIndex
			match = closingTagRegExp.exec(xml)
			if (match === null) {
				// `xml` is supposed to be "valid XML", so such situation isn't supposed to be possible.
				throw new Error(`Invalid XML: opening tag was found but closing tag was not: </${tagName}>`)
			}
			// Strictly speaking, this closing tag doesn't necessarily correspond to
			// the opening tag found previously. But because a precondition is that
			// `xml` is "valid XML", such erroneous situation isn't supposed to be possible.
			result.closingTagXml = match[0]
			result.closingTagStartIndex = match.index
			result.closingTagEndIndex = match.index + match[0].length - 1
			results.push(result)
		}
	}

	return results
}