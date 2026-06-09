// `fflate`'s `Zip` class is used to output a `.zip` archive in a "streaming" fashion.
//
// For compression, it uses `ZipDeflate` class, which enables DEFLATE compression algorithm.
// In the comments to `ZipDeflate` in `fflate` package, it says:
// "Streaming DEFLATE compression for ZIP archives. Prefer using AsyncZipDeflate for better performance"
//
import { Zip, ZipDeflate, AsyncZipDeflate } from 'fflate'

import { PassThrough } from 'node:stream'

// `fflate` author states that using `AsyncZipDeflate` instead of just `ZipDeflate`
// results in better performance because this way multiple files are compressed
// in-parallel rather than in-sequence. In other words, multiple CPU cores get utilized.
//
// But using this feature causes "DataCloneError: Cannot transfer object of unsupported type"
// error in Node.js
// https://github.com/101arrowz/fflate/issues/278
//
// Because of that, this flag is disabled.
//
const COMPRESS_FILES_IN_PARALLEL = false

/**
 * Creates a `*.zip` file from a map of files in a "streaming" fashion.
 * @param  {Record<string,Uint8Array>} files
 * @return {Stream} Returns a readable stream with the `.zip` archive data.
 */
export default function zipToStream(files) {
	// `stream` is a readable output stream with the `.zip` archive data.
	const stream = new PassThrough()

	// Tndividual file data sources that will be used by `fflate`'s `Zip` class instance.
	const fileDataSources = {}

	// `.zip` archive compression options.
	// By default, it uses `leveL: 6` of DEFLATE compression algorithm.
	const compressionOptions = undefined

	let errored = false

	const onError = (error) => {
		if (!errored) {
			errored = true
			// "To emit and properly handle errors in a Node.js `PassThrough` stream,
			//  you must explicitly call `stream.destroy(error)` instead of just invoking
			//  `.emit('error')`. Using `.destroy(error)` safely transitions the stream
			//  into an errored state, closes underlying resources, and guarantees that
			//  attached error handlers capture the failure."
			stream.destroy(error)
		}
	}

	// Starts writing `.zip` archive data into the `stream`.
	const start = (zip) => {
		// Push each file's data to be compressed by `fflate`'s `Zip` class instance.
		for (const filePath of Object.keys(fileDataSources)) {
			if (files[filePath] instanceof Uint8Array) {
				fileDataSources[filePath].push(files[filePath], true)
			} else {
				throw new Error(`Unsupported file data: ${files[filePath]}`)
			}
		}
		// Finalize the `.zip` archive:
		// no more files will be added to it.
		zip.end()
	}

	const zipperCallback = (error, dataChunk, isEndOfArchive) => {
		if (errored) {
			return
		}
		if (error) {
			onError(error)
			return
		}
		if (isEndOfArchive) {
			stream.end(dataChunk)
		} else {
			stream.write(dataChunk)
		}
	}

	// Create a `Zip` class instance.
	const zip = new Zip(zipperCallback)

	// Add the files to the archive.
	// Construct a file data source for each file.
	for (const filePath of Object.keys(files)) {
		const fileDataSource = COMPRESS_FILES_IN_PARALLEL
			? new AsyncZipDeflate(filePath, compressionOptions)
			: new ZipDeflate(filePath, compressionOptions)
		fileDataSources[filePath] = fileDataSource
		zip.add(fileDataSource)
	}

	// Start writing `.zip` archive data into the `stream`.
  start(zip)

	// Return a readable stream.
	return stream
}