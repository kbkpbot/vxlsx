module vxlsx

import time

#flag @VMODROOT/c/libxlsxwriter.o
#flag -lminizip
#flag -I @VMODROOT/c/include/xlsxwriter
#include "workbook.h"

pub type All_Data_Type = bool | f64 | int | string | time.Time

pub const null_format = unsafe { nil }

pub enum Lxw_error {
	lxw_no_error                                 = 0
	lxw_error_memory_malloc_failed
	lxw_error_creating_xlsx_file
	lxw_error_creating_tmpfile
	lxw_error_reading_tmpfile
	lxw_error_zip_file_operation
	lxw_error_zip_parameter_error
	lxw_error_zip_bad_zip_file
	lxw_error_zip_internal_error
	lxw_error_zip_file_add
	lxw_error_zip_close
	lxw_error_feature_not_supported
	lxw_error_null_parameter_ignored
	lxw_error_parameter_validation
	lxw_error_sheetname_length_exceeded
	lxw_error_invalid_sheetname_character
	lxw_error_sheetname_start_end_apostrophe
	lxw_error_sheetname_already_used
	lxw_error_32_string_length_exceeded
	lxw_error_128_string_length_exceeded
	lxw_error_255_string_length_exceeded
	lxw_error_max_string_length_exceeded
	lxw_error_shared_string_index_not_found
	lxw_error_worksheet_index_out_of_range
	lxw_error_worksheet_max_url_length_exceeded
	lxw_error_worksheet_max_number_urls_exceeded
	lxw_error_image_dimensions
	lxw_max_errno
}

fn C.lxw_version() &char
fn C.lxw_version_id() u16
fn C.lxw_strerror(error_num int) &char
fn C.lxw_datetime_to_excel_datetime(datetime &C.lxw_datetime) f64
fn C.lxw_unixtime_to_excel_date(unixtime i64) f64

// lxw_version retrieve the library version
pub fn lxw_version() string {
	return unsafe { tos_clone(&u8(C.lxw_version())) }
}

// lxw_version_id retrieve the library version ID
pub fn lxw_version_id() u16 {
	return C.lxw_version_id()
}

// lxw_strerror converts a libxlsxwriter error number to a string
pub fn lxw_strerror(error_num Lxw_error) string {
	return unsafe { tos_clone(&u8(C.lxw_strerror(int(error_num)))) }
}

// lxw_datetime_to_excel_datetime converts a lxw_datetime to an Excel datetime number
pub fn lxw_datetime_to_excel_datetime(value time.Time) f64 {
	datetime := &C.lxw_datetime{
		year: value.year
		month: value.month
		day: value.day
		hour: value.hour
		min: value.minute
		sec: value.second
	}
	return C.lxw_datetime_to_excel_datetime(datetime)
}

// lxw_unixtime_to_excel_date converts a unix datetime to an Excel datetime number
pub fn lxw_unixtime_to_excel_date(unixtime i64) f64 {
	return C.lxw_unixtime_to_excel_date(unixtime)
}
