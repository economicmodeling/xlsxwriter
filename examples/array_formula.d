/*
 * Example of how to use the libxlsxwriter library to write simple
 * array formulas.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */
import xlsxwriter;

void main() {
    /* Create a new workbook and add a worksheet. */
    lxw_workbook  *workbook  = workbook_new("array_formula.xlsx");
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, null);

    /* Write some data for the formulas. */
    worksheet_write_number(worksheet, 0, 1, 500, null);
    worksheet_write_number(worksheet, 1, 1, 10, null);
    worksheet_write_number(worksheet, 4, 1, 1, null);
    worksheet_write_number(worksheet, 5, 1, 2, null);
    worksheet_write_number(worksheet, 6, 1, 3, null);
    worksheet_write_number(worksheet, 0, 2, 300, null);
    worksheet_write_number(worksheet, 1, 2, 15, null);
    worksheet_write_number(worksheet, 4, 2, 20234, null);
    worksheet_write_number(worksheet, 5, 2, 21003, null);
    worksheet_write_number(worksheet, 6, 2, 10000, null);

    /* Write an array formula that returns a single value. */
    worksheet_write_array_formula(worksheet, 0, 0, 0, 0, "{=SUM(B1:C1*B2:C2)}", null);

    /* Similar to above but using the RANGE macro. */
    worksheet_write_array_formula(worksheet, RANGE("A2:A2").expand, "{=SUM(B1:C1*B2:C2)}", null);

    /* Write an array formula that returns a range of values. */
    worksheet_write_array_formula(worksheet, 4, 0, 6, 0, "{=TREND(C5:C7,B5:B7)}", null);
    workbook_close(workbook);
}
