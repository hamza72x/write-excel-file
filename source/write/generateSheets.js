import generateSheetXml from './files/sheet.xml/sheet.xml.js'
import initializeStyles from './initializeStyles.js'
import initializeSharedStrings from './initializeSharedStrings.js'
import validateSheetName from './validateSheetName.js'
import { getStylesParameters } from './parameters.js'

export default function generateSheets({
  sheetNames,
  headerStyle,
  getHeaderStyle,
  fontFamily,
  fontSize,
  features,
  ...restParameters
}) {
  const { getSharedStrings, getSharedString } = initializeSharedStrings()

  const { getStyles, getStyle } = initializeStyles({
    fontFamily,
    fontSize,
    features,
    ...getStylesParameters(restParameters, features)
  })

  // Rename deprecated `headerStyle` parameter to `getHeaderStyle(columnSchema)`.
  if (headerStyle && !getHeaderStyle) {
    getHeaderStyle = () => headerStyle
  }

  // Validate sheet name.
  for (const sheetName of sheetNames) {
    validateSheetName(sheetName)
  }

  const sheetsXml = []
  let sheetIndex = 0
  for (const sheetName of sheetNames) {
    const {
      data,
      schema,
      columns,
      images,
      ...sheetParameters
    } = restParameters
		const { multipleSheetsParameters } = restParameters
		validateData(data, { schema, multipleSheetsParameters })
		// Validate the cases when this function is called directly in tests, etc.
		if (multipleSheetsParameters === undefined) {
			throw new Error('`multipleSheetsParameters` parameter not found')
		}
		const getParameterForSheet = (parameter) => {
			if (multipleSheetsParameters) {
				return parameter[sheetIndex]
			}
			return parameter
		}
    sheetsXml.push(generateSheetXml(getParameterForSheet(data), {
      ...sheetParameters,
      sheetIndex,
      sheetId: sheetIndex + 1,
      schema: schema && getParameterForSheet(schema),
      columns: columns && getParameterForSheet(columns),
      images: images && getParameterForSheet(images),
      customFont: fontFamily || fontSize,
      getHeaderStyle,
      getStyle,
      getSharedString
    }, features))
    sheetIndex++
  }

  return {
    sheets: sheetNames.map((sheetName, i) => ({
      id: i + 1,
      name: sheetName,
      data: sheetsXml[i],
      images: restParameters.images && restParameters.images[i]
    })),
    getSharedStrings,
    getStyles
  }
}

function validateData(data, { schema, multipleSheetsParameters }) {
	if (multipleSheetsParameters) {
		if (!Array.isArray(data) || (data.length > 0 && !Array.isArray(data[0]))) {
			throw new TypeError('Expected `data` to be an array of arrays of rows')
		}
		for (const sheetData of data) {
			validateSheetDataRows(sheetData, { schema })
		}
	} else {
		if (!Array.isArray(data)) {
			throw new TypeError('Expected `data` to be an array of rows')
		}
		validateSheetDataRows(data, { schema })
	}
}

function validateSheetDataRows(data, { schema }) {
	for (const row of data) {
		if (row !== null && row !== undefined) {
			if (schema) {
				if (isObject(row)) {
					return
				} else {
					throw new Error('Expected `data` rows to be objects')
				}
			} else {
				if (Array.isArray(row)) {
					return
				} else {
					throw new Error('Expected `data` rows to be arrays')
				}
			}
		}
	}
}

const objectConstructor = {}.constructor

function isObject(object) {
  return object !== undefined && object !== null && object.constructor === objectConstructor
}
