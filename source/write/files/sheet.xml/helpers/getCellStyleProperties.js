import addProperties from '../../../helpers/addProperties.js'

export default function getCellStyleProperties(cell, features) {
	const {
		align,
		alignVertical,
		textRotation,
		indent,
		wrap,
		fontFamily,
		fontSize,
		fontWeight,
		fontStyle,
		textColor: textColorPropertyValue,
		// `color` property name is deprecated, use `textColor` property instead.
		color,
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
	} = cell

	// `color` property name is deprecated, use `textColor` property instead.
	const textColor = textColorPropertyValue || color

	// Add any custom cell style properties that're used by plugins.
	const additionalProperties = {}
	for (const feature of features) {
		if (feature.transforms && feature.transforms.style && feature.transforms.style.parameters) {
			addProperties(additionalProperties, cell, feature.transforms.style.parameters)
		}
	}

	if (
		align ||
		alignVertical ||
		textRotation ||
		indent ||
		wrap ||
		fontFamily ||
		fontSize ||
		fontWeight ||
		fontStyle ||
		textColor ||
		backgroundColor ||
		fillPatternStyle ||
		fillPatternColor ||
		borderColor ||
		borderStyle ||
		leftBorderColor ||
		leftBorderStyle ||
		rightBorderColor ||
		rightBorderStyle ||
		topBorderColor ||
		topBorderStyle ||
		bottomBorderColor ||
		bottomBorderStyle ||
		Object.keys(additionalProperties).length > 0
	) {
		return omitUndefinedProperties({
			...additionalProperties,
			align,
			alignVertical,
			textRotation,
			indent,
			wrap,
			fontFamily,
			fontSize,
			fontWeight,
			fontStyle,
			textColor,
			backgroundColor,
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
		})
	}
}

function omitUndefinedProperties(object) {
	const filteredObject = {}

	for (const key in object) {
		if (object[key] !== undefined) {
			filteredObject[key] = object[key]
		}
	}

	return filteredObject
}