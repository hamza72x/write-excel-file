// `Readable` type becomes globally available after installing `@types/node`.
// https://stackoverflow.com/questions/49508610/type-for-nodejs-stream-stream-in-typescript
// Or does it?
// https://gitlab.com/catamphetamine/write-excel-file/-/issues/53#note_1245377052

import Stream, { Readable } from 'stream'
import { PathLike } from 'fs';

import { ImageType } from '../types/features/images.d.js'

import {
	SheetData,
	OptionsSingleSheetWithSchema,
	OptionsMultipleSheetsWithSchema,
	OptionsSingleSheetWithoutSchema,
	OptionsMultipleSheetsWithoutSchema
} from './types/api.d.js'

export { ValueType, Cell, Row, SheetData, ColumnSchema, Schema, Feature } from '../types/api.d.js'

type ImageContent = Stream | Buffer | PathLike;

export type Image = ImageType<ImageContent>;

interface AdditionalOptionsReturnStream {}

interface AdditionalOptionsReturnBuffer {
	buffer: true;
}

interface AdditionalOptionsWriteToFile {
	filePath: string;
}

// export interface Stream {
// 	pipe(output: Stream): void;
// }

// If `filePath` parameter is not passed then the returned `Promise`
// resolves to a `Stream`-like object having a `.pipe()` method.

// With Schema.

interface OptionsSingleSheetWithSchemaWriteToFile<Object> extends OptionsSingleSheetWithSchema<Object, Image>, AdditionalOptionsWriteToFile<Object> {}
interface OptionsMultipleSheetsWithSchemaWriteToFile<Object> extends OptionsMultipleSheetsWithSchema<Object, Image>, AdditionalOptionsWriteToFile<Object> {}

interface OptionsSingleSheetWithSchemaReturnBuffer<Object> extends OptionsSingleSheetWithSchema<Object, Image>, AdditionalOptionsReturnBuffer<Object> {}
interface OptionsMultipleSheetsWithSchemaReturnBuffer<Object> extends OptionsMultipleSheetsWithSchema<Object, Image>, AdditionalOptionsReturnBuffer<Object> {}

interface OptionsSingleSheetWithSchemaReturnStream<Object> extends OptionsSingleSheetWithSchema<Object, Image>, AdditionalOptionsReturnStream<Object> {}
interface OptionsMultipleSheetsWithSchemaReturnStream<Object> extends OptionsMultipleSheetsWithSchema<Object, Image>, AdditionalOptionsReturnStream<Object> {}

// With `schema`.
// Write to a file.
declare function writeXlsxFile<Object>(
	objects: Object[] | Object[][],
	options: OptionsSingleSheetWithSchemaWriteToFile<Object> | OptionsMultipleSheetsWithSchemaWriteToFile<Object>
) : Promise<void>;

// With `schema`.
// Return `Buffer`.
declare function writeXlsxFile<Object>(
	objects: Object[] | Object[][],
	options: OptionsSingleSheetWithSchemaReturnBuffer<Object> | OptionsMultipleSheetsWithSchemaReturnBuffer<Object>
) : Promise<Buffer>;

// With `schema`.
// Return `Stream`.
declare function writeXlsxFile<Object>(
	objects: Object[] | Object[][],
	options: OptionsSingleSheetWithSchemaReturnStream<Object> | OptionsMultipleSheetsWithSchemaReturnStream<Object>
) : Promise<Readable>;

// Without Schema.

interface OptionsSingleSheetWithoutSchemaWriteToFile extends OptionsSingleSheetWithoutSchema<Image>, AdditionalOptionsWriteToFile {}
interface OptionsMultipleSheetsWithoutSchemaWriteToFile extends OptionsMultipleSheetsWithoutSchema<Image>, AdditionalOptionsWriteToFile {}

interface OptionsSingleSheetWithoutSchemaReturnBuffer extends OptionsSingleSheetWithoutSchema<Image>, AdditionalOptionsReturnBuffer {}
interface OptionsMultipleSheetsWithoutSchemaReturnBuffer extends OptionsMultipleSheetsWithoutSchema<Image>, AdditionalOptionsReturnBuffer {}

interface OptionsSingleSheetWithoutSchemaReturnStream extends OptionsSingleSheetWithoutSchema<Image>, AdditionalOptionsReturnStream {}
interface OptionsMultipleSheetsWithoutSchemaReturnStream extends OptionsMultipleSheetsWithoutSchema<Image>, AdditionalOptionsReturnStream {}

// Write to a file.
declare function writeXlsxFile(
	data: SheetData,
	options: OptionsSingleSheetWithoutSchemaWriteToFile
) : Promise<void>;

// Write to a file.
// Multiple `sheets`.
declare function writeXlsxFile(
	data: SheetData[],
	options: OptionsMultipleSheetsWithoutSchemaWriteToFile
) : Promise<void>;

// Return `Buffer`.
declare function writeXlsxFile(
	data: SheetData,
	options: OptionsSingleSheetWithoutSchemaReturnBuffer
) : Promise<Buffer>;

// Return `Buffer`.
// Multiple `sheets`.
declare function writeXlsxFile(
	data: SheetData[],
	options: OptionsMultipleSheetsWithoutSchemaReturnBuffer
) : Promise<Buffer>;

// Return `Stream`.
declare function writeXlsxFile(
	data: SheetData,
	options?: OptionsSingleSheetWithoutSchemaReturnStream
) : Promise<Readable>;

// Return `Stream`.
// Multiple `sheets`.
declare function writeXlsxFile(
	data: SheetData[],
	options?: OptionsMultipleSheetsWithoutSchemaReturnStream
) : Promise<Readable>;

export default writeXlsxFile;