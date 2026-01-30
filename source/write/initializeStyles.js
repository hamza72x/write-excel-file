import hasFont from './helpers/hasFont.js'
import hasFill from './helpers/hasFill.js'
import hasBorder from './helpers/hasBorder.js'

// There seem to be about 100 "built-in" formats in Excel.
// https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee857658(v=office.14)?redirectedfrom=MSDN
export const FORMAT_ID_STARTS_FROM = 100

export default function initializeStyles({
  fontFamily: defaultFontFamily,
  fontSize: defaultFontSize,
  features,
  ...restParameters
}) {
  const formats = []
  const formatIdByFormat = {}

  const styles = []
  const stylesIndex = {}

  const fonts = []
  const fontIdByFontKey = {}

  const fills = []
  const fillIdByFillKey = {}

  const borders = []
  const borderIdByBorderKey = {}

  // Default font.
  const defaultFont = {
    fontSize: defaultFontSize,
    fontFamily: defaultFontFamily
  }
  const defaultFontId = fonts.length
  fontIdByFontKey[getFontKey(defaultFont)] = defaultFontId
  fonts.push(defaultFont)

  // Default fill (no color).
  const defaultFill = {}
  fillIdByFillKey[getFillKey(defaultFill)] = fills.length
  fills.push(defaultFill)

  // Default border.
  const defaultBorder = {}
  borderIdByBorderKey[getBorderKey(defaultBorder)] = borders.length
  borders.push(defaultBorder)

  // Always add a mandatory "gray125" fill.
  // For some weird reason, MS Office 2007 Excel seems to require this type of `<fill>`
  // to always be present in an `*.xslx` document.
  // Otherwise, if this `<fill>` is absent in the document, it would forcefully overwrite
  // the first custom `backgroundColor` with a "gray125" fill.
  const gray125Fill = {
    gray125: true
  }
  fills.push(gray125Fill)

  // Returns a style ID number.
  // Adds a new style if it doesn't exist yet.
  // If it already exists, just reuses the existing one.
  function getStyleId({
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
    const font = {
      fontFamily,
      fontSize,
      fontWeight,
      fontStyle,
      textColor
    }

    const fill = {
      backgroundColor,
      fillPatternStyle,
      fillPatternColor
    }

    const border = {
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

    const fontKey = getFontKey(font)
    const fillKey = getFillKey(fill)
    const borderKey = getBorderKey(border)

    const getStyleKey = () => {
      return `${align || '-'}/${alignVertical || '-'}/${textRotation || '-'}/${indent || '-'}/${wrap || '-'}/${format || '-'}/${fontKey}/${fillKey}/${borderKey}`
    }

    // Returns style ID.
    const addStyle = () => {
      // Get format ID.
      let formatId
      if (format) {
        formatId = formatIdByFormat[format]
        if (formatId === undefined) {
          formatId = FORMAT_ID_STARTS_FROM + formats.length
          formatIdByFormat[format] = formatId
          formats.push(format)
        }
      }

      // Get font ID.
      let fontId
      if (hasFont(font)) {
        fontId = fontIdByFontKey[fontKey]
        if (fontId === undefined) {
          fontId = fonts.length
          fontIdByFontKey[fontKey] = fontId
          fonts.push({
            ...font,
            fontSize: font.fontSize || defaultFontSize,
            fontFamily: font.fontFamily || defaultFontFamily
          })
        }
      } else if (hasFont(defaultFont)) {
        fontId = defaultFontId
      }

      // Get fill ID.
      let fillId
      if (hasFill(fill)) {
        fillId = fillIdByFillKey[fillKey]
        if (fillId === undefined) {
          fillId = fills.length
          fillIdByFillKey[fillKey] = fillId
          fills.push({
            backgroundColor,
            fillPatternStyle,
            fillPatternColor
          })
        }
      }

      // Get border ID.
      let borderId
      if (hasBorder(border)) {
        borderId = borderIdByBorderKey[borderKey]
        if (borderId === undefined) {
          borderId = borders.length
          borderIdByBorderKey[borderKey] = borderId
          borders.push(border)
        }
      }

      // Apply any plugins that add custom styling.
      const customStyleProperties = {}
      for (const feature of features) {
        if (feature.transforms && feature.transforms.style) {
          customStyleProperties = feature.transforms.style.transform(customStyleProperties, restParameters)
        }
      }

      const styleId = styles.length
      stylesIndex[styleKey] = styleId

      // Add a style.
      styles.push({
        ...customStyleProperties,
        formatId,
        fontId,
        fillId,
        borderId,
        align,
        alignVertical,
        textRotation,
        indent,
        wrap
      })

      return styleId
    }

    let styleKey = getStyleKey()

    // Apply any plugins that add custom styling.
    for (const feature of features) {
      if (feature.transforms && feature.transforms.style) {
        styleKey = feature.transforms.style.transformKey(styleKey, restParameters)
      }
    }

    // Look for an existing style.
    if (stylesIndex[styleKey] !== undefined) {
      return stylesIndex[styleKey]
    }

    // Create new style if doesn't exist.
    return addStyle()
  }

  // Add the default style.
  getStyleId({}, {})

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
    getStyleId
  }
}

function getFontKey({
  fontFamily,
  fontSize,
  fontWeight,
  fontStyle,
  textColor
}) {
  return `${fontFamily || '-'}:${fontSize || '-'}:${fontWeight || '-'}:${fontStyle || '-'}:${textColor || '-'}`
}

function getFillKey({
  backgroundColor,
  fillPatternStyle,
  fillPatternColor
}) {
  return `${backgroundColor || '-'}:${fillPatternStyle || '-'}:${fillPatternColor || '-'}`
}

function getBorderKey({
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
}) {
  return `${(topBorderColor || borderColor) || '-'}:${(topBorderStyle || borderStyle) || '-'}` +
    '/' +
    `${(rightBorderColor || borderColor) || '-'}:${(rightBorderStyle || borderStyle) || '-'}` +
    '/' +
    `${(bottomBorderColor || borderColor) || '-'}:${(bottomBorderStyle || borderStyle) || '-'}` +
    '/' +
    `${(leftBorderColor || borderColor) || '-'}:${(leftBorderStyle || borderStyle) || '-'}`
}