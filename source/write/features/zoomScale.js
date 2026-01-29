// This feature is not included by default.
// It's included here just as an example of a feature implementation.
export default {
	transforms: {
		sheetXml: {
			transform: (xml, {}, { zoomScale }) => {
				if (typeof zoomScale === 'number') {
					// Excel 2007 doesn't like fractional `zoomScale` percentage values.
					// To work around that, it rounds it.
					return xml.replace(/<sheetView /g, `<sheetView zoomScale="${Math.round(zoomScale * 100)}" `)
				}
				return xml
			},
			// These parameters will be passed through to the `transform()` function above.
			parameters: ['zoomScale']
		}
	}
}

// Usage:
//
// await writeXlsxFile(data, {
// 	zoomScale: 1.5,
// 	features: [zoomScaleFeature],
// 	filePath: ...
// })
