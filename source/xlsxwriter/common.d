/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * common - Common functions and defines for the libxlsxwriter library.
 *
 */
module xlsxwriter.common;

extern(C):

/* enum lxw_boolean { */
/*     LXW_FALSE, */
/*     LXW_TRUE */
/* }; */

enum LXW_SHEETNAME_MAX =   32;
enum LXW_SHEETNAME_LEN =   65;
enum LXW_UINT32_T_LEN  =   11;   /* Length of 4294967296\0. */
enum LXW_IGNORE        =    1;
enum FILENAME_LEN      =  128;
enum LXW_NO_ERROR      =    0;

enum lxw_error {
	LXW_NO_ERROR, // No error.
	LXW_ERROR_MEMORY_MALLOC_FAILED, // Memory error, failed to malloc() required memory.
	LXW_ERROR_CREATING_XLSX_FILE, // Error creating output xlsx file. Usually a permissions error.
	LXW_ERROR_CREATING_TMPFILE, // Error encountered when creating a tmpfile during file assembly.
	LXW_ERROR_ZIP_FILE_OPERATION, // Zlib error with a file operation while creating xlsx file.
	LXW_ERROR_ZIP_FILE_ADD, // Zlib error when adding sub file to xlsx file.
	LXW_ERROR_ZIP_CLOSE, // Zlib error when closing xlsx file.
	LXW_ERROR_NULL_PARAMETER_IGNORED, // NULL function parameter ignored.
	LXW_ERROR_PARAMETER_VALIDATION, // Function parameter validation error.
	LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED, // String exceeds Excel's limit of 32,767 characters.
	LXW_ERROR_128_STRING_LENGTH_EXCEEDED, // Parameter exceeds Excel's limit of 128 characters.
	LXW_ERROR_255_STRING_LENGTH_EXCEEDED, // Parameter exceeds Excel's limit of 255 characters.
	LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND, // Error finding internal string index.
	LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, // Worksheet row or column index out of range.
	LXW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED, // Maximum number of worksheet URLs (65530) exceeded.
	LXW_ERROR_IMAGE_DIMENSIONS, // Couldn't read image dimensions or DPI.
}

alias lxw_row_t = uint;
alias lxw_col_t = ushort;

string standaloneEnums(alias enumName)()
{
	import std.traits : EnumMembers;
	enum e = __traits(identifier, enumName);

	import std.algorithm : map;
	import std.string : format, join;
	return [EnumMembers!enumName].map!(m => "enum %s = %s.%s;".format(m, e, m)).join();
}
