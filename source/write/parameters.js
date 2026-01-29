import addProperties from './helpers/addProperties.js'

export function getSheetsParameters(parameters, data, features) {
	const {
		schema,
		columns,
		images,
		sheet: sheetName,
		sheets: sheetNames,
		headerStyle,
		getHeaderStyle,
		fontFamily,
		fontSize,
		orientation,
		stickyRowsCount,
		stickyColumnsCount,
		showGridLines,
		rightToLeft,
		dateFormat
	} = parameters

	const sheetsParameters = {
		data,
		features,
		schema,
		columns,
		images,
		sheetNames,
		headerStyle,
		getHeaderStyle,
		fontFamily,
		fontSize,
		orientation,
		stickyRowsCount,
		stickyColumnsCount,
		showGridLines,
		rightToLeft,
		dateFormat
	}

	// Versions before `1.3.4` had a bug:
	// In a "write multiple sheets" scenario, `columns` parameter
	// wasn't required to be an array of `columns` for each sheet.
	if (sheetNames) {
		if (columns) {
			if (!Array.isArray(columns[0])) {
				throw new Error('In a "write multiple sheets" scenario, `columns` parameter must be an array of `columns` for each sheet.')
			}
		}
	}

	// Add any custom properties that're used by plugins.
	for (const feature of features) {
		if (feature.transforms && feature.transforms.sheetXml && feature.transforms.sheetXml.parameters) {
			addProperties(sheetsParameters, parameters, feature.transforms.sheetXml.parameters)
		}
	}

	// Also add any custom styles properties that're used by plugins
	// because styles properties are read from sheets properties.
	for (const feature of features) {
		if (feature.transforms && feature.transforms.stylesXml && feature.transforms.stylesXml.parameters) {
			addProperties(sheetsParameters, parameters, feature.transforms.stylesXml.parameters)
		}
	}

	// If only a single sheet is being written,
	// convert parameters to arrays as if multiple sheets were being written.
	// This way, the code after this wouldn't bother about the parameters being arrays or not.
	if (!sheetNames) {
		return {
			...sheetsParameters,
			multipleSheetsParameters: false,
			sheetNames: [sheetName || 'Sheet1']
		}
	}

	return {
		...sheetsParameters,
		multipleSheetsParameters: true
	}
}

export function getWorkbookParameters(parameters) {
	const {
		stickyRowsCount,
		stickyColumnsCount
	} = parameters

	return {
		stickyRowsCount,
		stickyColumnsCount
	}
}

export function getStylesParameters(parameters, features) {
	const stylesParameters = {}

	// Add any custom properties that're used by plugins.
	for (const feature of features) {
		if (feature.transforms && feature.transforms.stylesXml && feature.transforms.stylesXml.parameters) {
			addProperties(stylesParameters, parameters, feature.transforms.stylesXml.parameters)
		}
	}

	return stylesParameters
}

export function getContentTypesParameters(parameters) {
	const {
		images
	} = parameters

	return {
		images
	}
}