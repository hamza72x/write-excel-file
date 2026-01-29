import { Color, CommonStyleProperties } from './types/common.d.js'
import { ConditionalFormattingParametersSingleSheet, ConditionalFormattingParametersMultipleSheets } from './types/features/conditionalFormatting.d.js'

// Some users have requested exporting `ValueType` type.
// https://gitlab.com/catamphetamine/write-excel-file/-/issues/30
export type ValueType =
	String |
	Date |
	Number |
	Boolean;

// It's unclear how to express something like `type? = Type` in TypeScript.
// So instead it's defined as `type?: TypeConstructor<Type>`.
// https://gitlab.com/catamphetamine/write-excel-file/-/issues/4#note_715204034
// https://www.typescriptlang.org/docs/handbook/2/conditional-types.html
type TypeConstructor<Type> =
	Type extends String
		? StringConstructor
		: Type extends Date
			? DateConstructor
			: Type extends Number
				? NumberConstructor
				: Type extends Boolean
					? BooleanConstructor
					: never;

type TypeOfType<Type> = TypeConstructor<Type> | 'Formula';

interface CellStyle extends CommonStyleProperties {
	align?: 'left' | 'center' | 'right';
	alignVertical?: 'top' | 'center' | 'bottom';
	height?: number;
	span?: number;
	rowSpan?: number;
	indent?: number;
	wrap?: boolean;
	textRotation?: number;
	fontFamily?: string;
	fontSize?: number;
	/**
	 * @deprecated Use `textColor` property instead.
	 */
	color?: Color;
}

interface CellProps<Type> extends CellStyle {
	// TypeScript interprets `type?: Type` as "`type` is an instance of `Type`",
	// while in reality the definition of `type` is "`type` is `Type.constructor`".
	type?: TypeOfType<Type>;

	// A simpler (loose) variant:
	// type?: ValueType;

	// Data output format (for numbers or dates or strings).
	format?: string;
}

interface CellOfType<Type> extends CellProps<Type> {
	value?: Type;
}

export type Cell = CellOfType<ValueType> | null | undefined;
export type Row = Cell[];
export type SheetData = Row[];

// Some users have requested exporting `ColumnSchema` type.
// https://gitlab.com/catamphetamine/write-excel-file/-/issues/30
export interface ColumnSchema<Object, Type> extends CellProps<Type> {
	// Column title.
	column?: string;
	// Column width (in characters).
	width?: number;
	// Cell value getter.
	value(object: Object): Type | undefined | null;
	// Cell style getter.
	getCellStyle?(object: Object): CellStyle | undefined;
}

export type Schema<Object> = ColumnSchema<Object, ValueType>[];

type Column = {
	width?: number;
};

type Orientation = 'landscape';

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

type ImageContent = File | Blob | ArrayBuffer;

export type Image = ImageType<ImageContent>;

export type Columns = Column[];

export interface CommonOptions<Object = never> {
  // `headerStyle` parameter is deprecated, use `getHeaderStyle(columnSchema)` instead.
  headerStyle?: CellStyle;
  getHeaderStyle?: (columnSchema: ColumnSchema<Object, ValueType>) => CellStyle | undefined;
  fontFamily?: string;
  fontSize?: number;
  orientation?: Orientation;
  stickyColumnsCount?: number;
  stickyRowsCount?: number;
  showGridLines?: boolean;
  rightToLeft?: boolean;
  dateFormat?: string;
}

interface CommonOptionsWithoutFileName<Object = never> extends CommonOptions<Object> {}

interface CommonOptionsWriteToFile<Object = never> extends CommonOptions<Object> {
	fileName: string;
}

export interface Feature {
	transforms?: {
		sheetXml?: {
			transform: (
				xml: string,
				options: { sheetId: number, multipleSheetsParameters: boolean },
				customTopLevelParameters: Record<string, any>
			) => string;

			// These parameters will be passed through to the `transform()` function above.
			parameters?: string[];
		},

		stylesXml?: {
			transform: (
				xml: string,
				options: {},
				customTopLevelParameters: Record<string, any>
			) => string;

			// These parameters will be passed through to the `transform()` function above.
			parameters?: string[];
		},

		style?: {
			transform: (
				properties: Record<string, any>,
				customCellStyleParameters: Record<string, any>
			) => Record<string, any>;

			transformKey: (
				key: string,
				customCellStyleParameters: Record<string, any>
			) => string;

			// These parameters will be passed through to the `transform()` and `transformKey()` functions above.
			parameters?: string[];
		}
	};
}

interface CommonOptions {
	features?: Feature[];
}

interface CommonOptionsSingleSheet<Image> extends CommonOptions, ConditionalFormattingParametersSingleSheet {
  images?: Image[];
	sheet?: string;
}

interface CommonOptionsMultipleSheets<Image> extends CommonOptions, ConditionalFormattingParametersMultipleSheets {
  images?: Image[][];
	sheets?: string[];
}

// With Schema.

export interface OptionsWithSchemaAndSingleSheet<Image> extends CommonOptionsSingleSheet<Image> {
	schema: Schema<Object>;
}

export interface OptionsWithSchemaAndMultipleSheets<Image> extends CommonOptionsMultipleSheets<Image> {
	schema: Schema<Object>[];
}

interface OptionsWithSchemaWithoutFileNameSingleSheet<Object> extends CommonOptionsWithoutFileName<Object>, OptionsWithSchemaAndSingleSheet<Image> {
}

interface OptionsWithSchemaWriteToFileSingleSheet<Object> extends CommonOptionsWriteToFile<Object>, OptionsWithSchemaAndSingleSheet<Image> {
}

interface OptionsWithSchemaWithoutFileNameMultipleSheets<Object> extends CommonOptionsWithoutFileName<Object>, OptionsWithSchemaAndMultipleSheets<Image> {
}

interface OptionsWithSchemaWriteToFileMultipleSheets<Object> extends CommonOptionsWriteToFile<Object>, OptionsWithSchemaAndMultipleSheets<Image> {
}

// With `schema`.
// Write to a file.
declare function writeXlsxFile<Object>(
	objects: Object[],
	options: OptionsWithSchemaWriteToFileSingleSheet<Object>
) : Promise<void>;

// With `schema`.
// Return `Blob`.
declare function writeXlsxFile<Object>(
	objects: Object[],
	options: OptionsWithSchemaWithoutFileNameSingleSheet<Object>
) : Promise<Blob>;

// With `schema`.
// Write to a file.
// Multiple `sheets`.
declare function writeXlsxFile<Object>(
	objects: Object[][],
	options: OptionsWithSchemaWriteToFileMultipleSheets<Object>
) : Promise<void>;

// With `schema`.
// Return `Blob`.
// Multiple `sheets`.
declare function writeXlsxFile<Object>(
	objects: Object[][],
	options: OptionsWithSchemaWithoutFileNameMultipleSheets<Object>
) : Promise<Blob>;

// Without Schema.

export interface OptionsWithoutSchemaSingleSheet<Image> extends CommonOptionsSingleSheet<Image> {
	columns?: Columns;
}

export interface OptionsWithoutSchemaMultipleSheets<Image> extends CommonOptionsMultipleSheets<Image> {
	columns?: Columns[];
}

interface OptionsWithoutSchemaWithoutFileNameSingleSheet extends CommonOptionsWithoutFileName, OptionsWithoutSchemaSingleSheet<Image> {
}

interface OptionsWithoutSchemaWriteToFileSingleSheet extends CommonOptionsWriteToFile, OptionsWithoutSchemaSingleSheet<Image> {
}

interface OptionsWithoutSchemaWithoutFileNameMultipleSheets extends CommonOptionsWithoutFileName, OptionsWithoutSchemaMultipleSheets<Image> {
}

interface OptionsWithoutSchemaWriteToFileMultipleSheets extends CommonOptionsWriteToFile, OptionsWithoutSchemaMultipleSheets<Image> {
}

// Return `Blob`.
declare function writeXlsxFile(
	data: SheetData,
	options: OptionsWithoutSchemaWithoutFileNameSingleSheet
) : Promise<Blob>;

// Return `Blob`.
// Multiple `sheets`.
declare function writeXlsxFile(
	data: SheetData[],
	options: OptionsWithoutSchemaWithoutFileNameMultipleSheets
) : Promise<Blob>;

// Wrtie to a file.
declare function writeXlsxFile(
	data: SheetData,
	options: OptionsWithoutSchemaWriteToFileSingleSheet
) : Promise<void>;

// Wrtie to a file.
// Multiple `sheets`.
declare function writeXlsxFile(
	data: SheetData[],
	options: OptionsWithoutSchemaWriteToFileMultipleSheets
) : Promise<void>;

export default writeXlsxFile;