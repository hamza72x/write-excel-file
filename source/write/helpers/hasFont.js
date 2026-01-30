export default function hasFont({
	fontFamily,
	fontSize,
	fontWeight,
	fontStyle,
	textDecoration,
	textColor
}) {
	return Boolean(
		fontFamily ||
		typeof fontSize === 'number' ||
		fontWeight ||
		fontStyle ||
		textDecoration ||
		textColor
	)
}