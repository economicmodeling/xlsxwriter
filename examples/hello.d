/*
 * Example of writing some data to a simple Excel file using libxlsxwriter.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 * Translated to D 2016, Justin Whear
 */

import std.exception : enforce;
import xlsxwriter;

void main(string[] args)
{
	auto workbook = enforce( workbook_new("hello_world.xlsx") );
	auto worksheet = enforce( workbook_add_worksheet(workbook, null) );

	worksheet_write_string(worksheet, 0, 0, "Hello", null);
	worksheet_write_number(worksheet, 1, 0, 123, null);

	workbook_close(workbook);
}
