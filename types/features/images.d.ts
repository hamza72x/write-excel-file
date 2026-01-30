export interface ImagesParametersSingleSheet<Image> {
  images?: Image[];
}

export interface ImagesParametersMultipleSheets<Image> {
  images?: Image[][];
}

// `ImageContent` will differ between server side and client side.
export interface ImageType<ImageContent> {
	content: ImageContent;
	contentType: string;
	width: number;
	height: number;
	dpi: number;
	anchor: {
		row: number;
		column: number;
	};
	offsetX?: number;
	offsetY?: number;
	title?: string;
	description?: string;
}
