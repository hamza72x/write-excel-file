import escapeXmlCharacters from '../../xml/escapeXmlCharacters.js'
import $attr from '../../xml/sanitizeAttributeValue.js'

import { hasFill, getFillXml } from '../helpers/fill.js'
import { getBorderXml } from '../helpers/border.js'
import getXlsxColorForHexColor from '../helpers/getXlsxColorForHexColor.js'
import getCellCoordinate from '../helpers/getCellCoordinate.js'

export default {
	transforms: {
		sheetXml: {
			transform: (xml, { sheetIndex, multipleSheetsParameters }, { conditionalFormatting: conditionalFormattingParameter }) => {
				if (conditionalFormattingParameter) {
					const conditionalFormattingPerSheet = multipleSheetsParameters ? conditionalFormattingParameter : [conditionalFormattingParameter]
					const conditionalFormatting = multipleSheetsParameters ? conditionalFormattingParameter[sheetIndex] : conditionalFormattingParameter
					if (conditionalFormatting) {
						const conditionalFormattingRulesXml = getConditionalFormattingRulesXml({ sheetIndex, conditionalFormatting, conditionalFormattingPerSheet })
						if (conditionalFormattingRulesXml) {
							return xml.replace(/<\/worksheet>$/, conditionalFormattingRulesXml + '</worksheet>')
						}
					}
				}
				return xml
			},

			// These parameters will be passed through to the `transform()` function above.
			parameters: ['conditionalFormatting']
		},

		stylesXml: {
			transform: (xml, { multipleSheetsParameters }, { conditionalFormatting: conditionalFormattingParameter }) => {
				if (conditionalFormattingParameter) {
					const conditionalFormattingPerSheet = multipleSheetsParameters ? conditionalFormattingParameter : [conditionalFormattingParameter]
					const conditionalFormattingStylesXml = getConditionalFormattingStylesXml({ conditionalFormattingPerSheet })
					if (conditionalFormattingStylesXml) {
						return xml.replace(/<\/styleSheet>$/, conditionalFormattingStylesXml + '</styleSheet>')
					}
				}
				return xml
			},

			// These parameters will be passed through to the `transform()` function above.
			parameters: ['conditionalFormatting']
		}
	}
}

function getConditionalFormattingRulesXml({
	sheetIndex,
	conditionalFormatting: conditionalFormattingRules,
	conditionalFormattingPerSheet
}) {
	let xml = ''

	let i = 0
	for (const conditionalFormattingRule of conditionalFormattingRules) {
		const {
			cells: {
				from,
				to
			},
			condition: {
				formula,
				operator,
				value,
				value2
			}
		} = conditionalFormattingRule

		// Cell range example: "A1:C5"
		const cellRange = getCellCoordinate(from.row - 1, from.column - 1) + ':' + getCellCoordinate(to.row - 1, to.column - 1)

		xml += `<conditionalFormatting sqref="${cellRange}">`

		// The priority attribute in conditional formatting XML is an integer value that must be a positive integer between 1 and the total number of conditional formatting rules on the worksheet.
		// * The value is defined by the W3C XML Schema int datatype.
		// * The valid range is dynamic, starting from 1 (highest priority) and ending at the total number of rules on that specific worksheet (lowest priority).
		// * Each priority value must be unique for all rules on the same worksheet. Changing the priority of one rule will automatically adjust the priorities of other rules to maintain uniqueness and order.
		// * A lower numeric value indicates a higher priority (e.g., priority 1 is the highest possible priority).
		const priority = i + 1

		// The `dxfId` attribute in ".xlsx" (OpenXML) files is a zero-based integer index (0, 1, 2, ...)
		// referencing a specific differential format record (`<dxf>`) within the "styles.xml" file.
		const dxfId = getConditionalFormattingRuleGlobalIndex({
			conditionalFormattingRuleIndex: i,
			conditionalFormattingPerSheet,
			sheetIndex
		})

		// There're a lot of possible `type`s of a `<cfRule/>`.
		// The full list could be viewed by googling for "ST_CfType".

		if (formula) {
			xml += `<cfRule type="expression" dxfId="${dxfId}" priority="${priority}">`
			xml += `<formula>${escapeXmlCharacters(formula, { attribute: false })}</formula>`
			xml += '</cfRule>'
		} else if (operator) {
			xml += `<cfRule type="cellIs" operator="${getXlsxOperatorName(operator)}" dxfId="${dxfId}" priority="${priority}">`
			xml += `<formula>${escapeXmlCharacters(formatValue(value), { attribute: false })}</formula>`
			if (getXlsxOperatorName(operator) === 'between') {
				xml += `<formula>${escapeXmlCharacters(formatValue(value2), { attribute: false })}</formula>`
			}
			xml += '</cfRule>'
		} else {
			throw new Error(`Invalid conditional formatting rule:\n${JSON.stringify(conditionalFormattingRule, null, 2)}`)
		}

		xml += '</conditionalFormatting>'

		i++
	}

	return xml
}

function formatValue(value) {
	if (typeof value === 'string') {
		return '"' + value + '"'
	}
	return String(value)
}

function getConditionalFormattingRuleGlobalIndex({
	conditionalFormattingRuleIndex,
	conditionalFormattingPerSheet,
	sheetIndex
}) {
	let sheetRulesIndexOffset = 0

	let i = 0
	while (i < sheetIndex) {
		if (conditionalFormattingPerSheet[i]) {
			sheetRulesIndexOffset += conditionalFormattingPerSheet[i].length
		}
		i++
	}

	return sheetRulesIndexOffset + conditionalFormattingRuleIndex
}

function getXlsxOperatorName(operator) {
	switch (operator) {
		case '<':
			return 'lessThan'
		case '>':
			return 'greaterThan'
		case '<=':
			return 'lessThanOrEqual'
		case '>=':
			return 'greaterThanOrEqual'
		case '=':
			return 'equal'
		case '!=':
			return 'notEqual'
		case '...':
			return 'between'
		default:
			throw new Error(`Unknown conditional formatting operator: ${operator}`)
	}
}

function getConditionalFormattingStylesXml({ conditionalFormattingPerSheet }) {
	let totalConditionalFormattingRulesCount = 0

	for (const conditionalFormattingOfSheet of conditionalFormattingPerSheet) {
		totalConditionalFormattingRulesCount += conditionalFormattingOfSheet.length
	}

	let xml = ''

	xml += `<dxfs count="${totalConditionalFormattingRulesCount}">`

	for (const conditionalFormattingOfSheet of conditionalFormattingPerSheet) {
		for (const conditionalFormattingRule of conditionalFormattingOfSheet) {
			const {
				style: {
					fontWeight,
					fontStyle,
					textColor,
					textDecoration,
					backgroundColor,
					fillPatternStyle,
					fillPatternColor,
					borderColor,
					borderStyle,
					leftBorderColor,
					leftBorderStyle,
					rightBorderColor,
					rightBorderStyle,
					topBorderColor,
					topBorderStyle,
					bottomBorderColor,
					bottomBorderStyle
				}
			} = conditionalFormattingRule

			xml += '<dxf>'

			if (fontWeight || fontStyle || textColor || textDecoration) {
				xml += '<font>'

				if (fontWeight === 'bold') {
					xml += '<b/>'
				} else if (fontStyle === 'italic') {
					xml += '<i/>'
				} else if (textDecoration === 'strikethrough') {
					xml += '<strike/>'
				} else if (textDecoration === 'underline') {
					xml += '<u/>'
				} else if (textDecoration === 'double-underline') {
					xml += '<u val="double"/>'
				}

				if (textColor) {
					xml += `<color rgb="${$attr(getXlsxColorForHexColor(textColor))}"/>`
				}

				xml += '</font>'
			}

			const fill = {
				backgroundColor,
				fillPatternStyle,
				fillPatternColor
			}
			if (hasFill(fill)) {
				xml += getFillXml(fill)
			}

			if (
				borderColor ||
				borderStyle ||
				leftBorderColor ||
				leftBorderStyle ||
				rightBorderColor ||
				rightBorderStyle ||
				topBorderColor ||
				topBorderStyle ||
				bottomBorderColor ||
				bottomBorderStyle
			) {
				xml += getBorderXml({
          left: {
            style: leftBorderStyle || borderStyle,
            color: leftBorderColor || borderColor
          },
          right: {
            style: rightBorderStyle || borderStyle,
            color: rightBorderColor || borderColor
          },
          top: {
            style: topBorderStyle || borderStyle,
            color: topBorderColor || borderColor
          },
          bottom: {
            style: bottomBorderStyle || borderStyle,
            color: bottomBorderColor || borderColor
          }
				})
			}

			xml += '</dxf>'
		}
	}

	xml += '</dxfs>'

	return xml
}