import { ImageType } from './types/features/images.d.js'

import {
	SheetData,
	OptionsSingleSheetWithSchema,
	OptionsMultipleSheetsWithSchema,
	OptionsSingleSheetWithoutSchema,
	OptionsMultipleSheetsWithoutSchema
} from './types/api.d.js'

export { ValueType, Cell, Row, SheetData, ColumnSchema, Schema, Feature } from './types/api.d.js'

type ImageContent = File | Blob | ArrayBuffer;

export type Image = ImageType<ImageContent>;

interface AdditionalOptionsReturnBlob {}

interface AdditionalOptionsDownloadFile {
	fileName: string;
}

// With Schema.

interface OptionsSingleSheetWithSchemaReturnBlob<Object> extends OptionsSingleSheetWithSchema<Object, Image>, AdditionalOptionsReturnBlob {}
interface OptionsMultipleSheetsWithSchemaReturnBlob<Object> extends OptionsMultipleSheetsWithSchema<Object, Image>, AdditionalOptionsReturnBlob {}

interface OptionsSingleSheetWithSchemaDownloadFile<Object> extends OptionsSingleSheetWithSchema<Object, Image>, AdditionalOptionsDownloadFile {}
interface OptionsMultipleSheetsWithSchemaDownloadFile<Object> extends OptionsMultipleSheetsWithSchema<Object, Image>, AdditionalOptionsDownloadFile {}

// With `schema`.
// Download file.
declare function writeXlsxFile<Object>(
	objects: Object[],
	options: OptionsSingleSheetWithSchemaDownloadFile<Object>
) : Promise<void>;

// With `schema`.
// Return `Blob`.
declare function writeXlsxFile<Object>(
	objects: Object[],
	options: OptionsSingleSheetWithSchemaReturnBlob<Object>
) : Promise<Blob>;

// With `schema`.
// Download file.
// Multiple `sheets`.
declare function writeXlsxFile<Object>(
	objects: Object[][],
	options: OptionsMultipleSheetsWithSchemaDownloadFile<Object>
) : Promise<void>;

// With `schema`.
// Return `Blob`.
// Multiple `sheets`.
declare function writeXlsxFile<Object>(
	objects: Object[][],
	options: OptionsMultipleSheetsWithSchemaReturnBlob<Object>
) : Promise<Blob>;

// Without Schema.

interface OptionsSingleSheetWithoutSchemaReturnBlob extends OptionsSingleSheetWithoutSchema<Image>, AdditionalOptionsReturnBlob {}
interface OptionsMultipleSheetsWithoutSchemaReturnBlob extends OptionsMultipleSheetsWithoutSchema<Image>, AdditionalOptionsReturnBlob {}

interface OptionsSingleSheetWithoutSchemaDownloadFile extends OptionsSingleSheetWithoutSchema<Image>, AdditionalOptionsDownloadFile {}
interface OptionsMultipleSheetsWithoutSchemaDownloadFile extends OptionsMultipleSheetsWithoutSchema<Image>, AdditionalOptionsDownloadFile {}

// Return `Blob`.
declare function writeXlsxFile(
	data: SheetData,
	options: OptionsSingleSheetWithoutSchemaReturnBlob
) : Promise<Blob>;

// Return `Blob`.
// Multiple `sheets`.
declare function writeXlsxFile(
	data: SheetData[],
	options: OptionsMultipleSheetsWithoutSchemaReturnBlob
) : Promise<Blob>;

// Download file.
declare function writeXlsxFile(
	data: SheetData,
	options: OptionsSingleSheetWithoutSchemaDownloadFile
) : Promise<void>;

// Download file.
// Multiple `sheets`.
declare function writeXlsxFile(
	data: SheetData[],
	options: OptionsMultipleSheetsWithoutSchemaDownloadFile
) : Promise<void>;

export default writeXlsxFile;