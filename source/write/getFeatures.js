import conditionalFormatting from './features/conditionalFormatting.js'

export default function getFeatures(customFeatures) {
	const features = customFeatures ? customFeatures.slice() : []

	// Add "conditional formatting" feature implicitly.
	if (features.indexOf(conditionalFormatting) < 0) {
		features.push(conditionalFormatting)
	}

	return features
}