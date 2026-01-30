import { Color, CommonStyleProperties } from './common.d.js'

import { ConditionalFormattingParametersSingleSheet, ConditionalFormattingParametersMultipleSheets } from './features/conditionalFormatting.d.js'
import { ImagesParametersSingleSheet, ImagesParametersMultipleSheets } from './features/images.js'

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

export type Columns = Column[];

export interface CommonOptions {
	fontFamily?: string;
	fontSize?: number;
	orientation?: Orientation;
	stickyColumnsCount?: number;
	stickyRowsCount?: number;
	showGridLines?: boolean;
	rightToLeft?: boolean;
	dateFormat?: string;
	features?: Feature[];
}

export interface AdditionalOptionsWhenUsingSchema<Object = never> {
	// `headerStyle` parameter is deprecated, use `getHeaderStyle(columnSchema)` instead.
	headerStyle?: CellStyle;
	getHeaderStyle?: (columnSchema: ColumnSchema<Object, ValueType>) => CellStyle | undefined;
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

// Options.

interface CommonOptionsSingleSheet<Image> extends CommonOptions, ConditionalFormattingParametersSingleSheet, ImagesParametersSingleSheet<Image> {
	sheet?: string;
}

interface CommonOptionsMultipleSheets<Image> extends CommonOptions, ConditionalFormattingParametersMultipleSheets, ImagesParametersMultipleSheets<Image> {
	sheets?: string[];
}

// With Schema.

export interface OptionsSingleSheetWithSchema<Object, Image> extends CommonOptionsSingleSheet<Image>, AdditionalOptionsWhenUsingSchema<Object> {
	schema: Schema<Object>;
}

export interface OptionsMultipleSheetsWithSchema<Object, Image> extends CommonOptionsMultipleSheets<Image>, AdditionalOptionsWhenUsingSchema<Object> {
	schema: Schema<Object>[];
}

// Without Schema.

export interface OptionsSingleSheetWithoutSchema<Image> extends CommonOptionsSingleSheet<Image> {
	columns?: Columns;
}

export interface OptionsMultipleSheetsWithoutSchema<Image> extends CommonOptionsMultipleSheets<Image> {
	columns?: Columns[];
}
