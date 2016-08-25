/*
 * A simple program to write some data to an Excel file using the
 * libxlsxwriter library.
 *
 * This program is shown, with explanations, in Tutorial 3 of the
 * libxlsxwriter documentation.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */
import xlsxwriter;

/* Some data we want to write to the worksheet. */
struct expense {
    char[32]     item;
    int          cost;
    lxw_datetime datetime;
}

expense[] expenses = [
    {"Rent", 1000, { year: 2013, month: 1, day: 13 } },
    {"Gas",   100, { year: 2013, month: 1, day: 14 } },
    {"Food",  300, { year: 2013, month: 1, day: 16 } },
    {"Gym",    50, { year: 2013, month: 1, day: 20 } },
];

int main() {
    /* Create a workbook and add a worksheet. */
    lxw_workbook  *workbook  = workbook_new("tutorial03.xlsx");
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, null);
    lxw_row_t row = 0;
    lxw_col_t col = 0;
    int i;

    /* Add a bold format to use to highlight cells. */
    lxw_format *bold = workbook_add_format(workbook);
    format_set_bold(bold);

    /* Add a number format for cells with money. */
    lxw_format *money = workbook_add_format(workbook);
    format_set_num_format(money, "$#,##0");

    /* Add an Excel date format. */
    lxw_format *date_format = workbook_add_format(workbook);
    format_set_num_format(date_format, "mmmm d yyyy");

    /* Adjust the column width. */
    worksheet_set_column(worksheet, 0, 0, 15, null);

    /* Write some data header. */
    worksheet_write_string(worksheet, row, col, "Item", bold);
    worksheet_write_string(worksheet, row, cast(lxw_col_t)(col + 1), "Cost", bold);

    /* Iterate over the data and write it out element by element. */
    for (i = 0; i < 4; i++) {
        /* Write from the first cell below the headers. */
        row = i + 1;
        worksheet_write_string(worksheet, row, col, expenses[i].item.ptr, null);
        assert(0 == worksheet_write_datetime(worksheet, row, cast(lxw_col_t)(col + 1),
		                         &expenses[i].datetime, date_format));
        worksheet_write_number(worksheet, row, cast(lxw_col_t)(col + 2),
		                       expenses[i].cost, money);
    }

    /* Write a total using a formula. */
    worksheet_write_string (worksheet, row + 1, col,     "Total",       bold);
    worksheet_write_formula(worksheet, row + 1, cast(lxw_col_t)(col + 2), "=SUM(C2:C5)", money);

    /* Save the workbook and free any allocated memory. */
    return workbook_close(workbook);
}
