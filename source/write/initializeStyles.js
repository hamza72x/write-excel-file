// There seem to be about 100 "built-in" formats in Excel.
// https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee857658(v=office.14)?redirectedfrom=MSDN
export const FORMAT_ID_STARTS_FROM = 100

export default function initializeStyles({
  fontFamily: defaultFontFamily,
  fontSize: defaultFontSize,
  features,
  ...restParameters
}) {
  const customFont = Boolean(defaultFontFamily || defaultFontSize)

  if (defaultFontFamily === undefined) {
    defaultFontFamily = 'Calibri'
  }

  if (defaultFontSize === undefined) {
    defaultFontSize = 12
  }

  const formats = []
  const formatsIndex = {}

  const styles = []
  const stylesIndex = {}

  const fonts = []
  const fontsIndex = {}

  const fills = []
  const fillsIndex = {}

  const borders = []
  const bordersIndex = {}

  // Default font.
  fonts.push({
    size: defaultFontSize,
    family: defaultFontFamily,
    custom: customFont
  })
  fontsIndex['-:-'] = 0

  // Default fill (no color).
  fills.push({})
  fillsIndex['-'] = 0

  // Default border.
  borders.push({
    left: {},
    right: {},
    top: {},
    bottom: {}
  })
  bordersIndex['-:-/-:-/-:-/-:-'] = 0

  // Always add a mandatory "gray125" fill.
  // For some weird reason, MS Office 2007 Excel seems to require this type of `<fill>`
  // to always be present in an `*.xslx` document.
  // Otherwise, if this `<fill>` is absent in the document, it would forcefully overwrite
  // the first custom `backgroundColor` with a "gray125" fill.
  fills.push({
    gray125: true
  })

  function getStyle({
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
    bottomBorderStyle,
    ...restParameters
  }, {
    format
  }) {
    // Look for an existing style.
    const fontKey = `${fontFamily || '-'}:${fontSize || '-'}:${fontWeight || '-'}:${fontStyle || '-'}:${textColor || '-'}`
    const fillKey = `${backgroundColor || '-'}:${fillPatternStyle || '-'}:${fillPatternColor || '-'}`
    const borderKey =
      `${(topBorderColor || borderColor) || '-'}:${(topBorderStyle || borderStyle) || '-'}` +
      '/' +
      `${(rightBorderColor || borderColor) || '-'}:${(rightBorderStyle || borderStyle) || '-'}` +
      '/' +
      `${(bottomBorderColor || borderColor) || '-'}:${(bottomBorderStyle || borderStyle) || '-'}` +
      '/' +
      `${(leftBorderColor || borderColor) || '-'}:${(leftBorderStyle || borderStyle) || '-'}`

    let key = `${align || '-'}/${alignVertical || '-'}/${textRotation || '-'}/${indent || '-'}/${wrap || '-'}/${format || '-'}/${fontKey}/${fillKey}/${borderKey}`

    // Apply any plugins that add custom styling.
    for (const feature of features) {
      if (feature.transforms && feature.transforms.style) {
        key = feature.transforms.style.transformKey(key, restParameters)
      }
    }

    const styleId = stylesIndex[key]
    if (styleId !== undefined) {
      return styleId
    }

    // Create new style.

    // Get format ID.
    let formatId
    if (format) {
      formatId = formatsIndex[format]
      if (formatId === undefined) {
        formatId = formatsIndex[format] = String(FORMAT_ID_STARTS_FROM + formats.length)
        formats.push(format)
      }
    }

    // Get font ID.
    let fontId = customFont ? 0 : undefined
    if (fontFamily || fontSize || fontWeight || fontStyle || textColor) {
      fontId = fontsIndex[fontKey]
      if (fontId === undefined) {
        fontId = fontsIndex[fontKey] = String(fonts.length)
        fonts.push({
          custom: true,
          size: fontSize || defaultFontSize,
          family: fontFamily || defaultFontFamily,
          weight: fontWeight,
          style: fontStyle,
          color: textColor
        })
      }
    }

    // Get fill ID.
    let fillId
    if (backgroundColor || fillPatternStyle || fillPatternColor) {
      fillId = fillsIndex[fillKey]
      if (fillId === undefined) {
        fillId = fillsIndex[fillKey] = String(fills.length)
        fills.push({
          backgroundColor,
          fillPatternStyle,
          fillPatternColor
        })
      }
    }

    // Get border ID.
    let borderId
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
      borderId = bordersIndex[borderKey]
      if (borderId === undefined) {
        borderId = bordersIndex[borderKey] = String(borders.length)
        borders.push({
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
    }

    // Apply any plugins that add custom styling.
    const customStyleProperties = {}
    for (const feature of features) {
      if (feature.transforms && feature.transforms.style) {
        customStyleProperties = feature.transforms.style.transform(customStyleProperties, restParameters)
      }
    }

    // Add a style.
    styles.push({
      ...customStyleProperties,
      fontId,
      fillId,
      borderId,
      align,
      alignVertical,
      textRotation,
      indent,
      wrap,
      formatId
    })

    return stylesIndex[key] = String(styles.length - 1)
  }

  // Add the default style.
  getStyle({}, {})

  return {
    getStyles() {
      return {
        formats,
        styles,
        fonts,
        fills,
        borders,
        ...restParameters
      }
    },
    getStyle
  }
}