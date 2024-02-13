module vxlsx

import time

@[typedef]
pub struct C.lxw_doc_properties {
pub mut:
	title          &char
	subject        &char
	author         &char
	manager        &char
	company        &char
	category       &char
	keywords       &char
	comments       &char
	status         &char
	hyperlink_base &char
	created        C.time_t
}

pub struct WorkBook_Options {
pub mut:
	constant_memory u8
	tmpdir          string
	use_zip64       u8
	output_buffer   []u8
}

@[typedef]
pub struct C.lxw_workbook_options {
pub mut:
	constant_memory    u8
	tmpdir             &char
	use_zip64          u8
	output_buffer      &&char
	output_buffer_size &usize
}

@[typedef]
struct C.lxw_workbook {}

// ===========workbook functions===========
fn C.workbook_new(filename &char) &C.lxw_workbook
fn C.workbook_new_opt(filename &char, options &C.lxw_workbook_options) &C.lxw_workbook
fn C.workbook_add_worksheet(workbook &C.lxw_workbook, sheetname &char) &C.lxw_worksheet
fn C.workbook_add_chartsheet(workbook &C.lxw_workbook, sheetname &char) &C.lxw_chartsheet
fn C.workbook_add_format(workbook &C.lxw_workbook) &C.lxw_format
fn C.workbook_add_chart(workbook &C.lxw_workbook, chart_type u8) &C.lxw_chart
fn C.workbook_close(workbook &C.lxw_workbook) Lxw_error
fn C.workbook_set_properties(workbook &C.lxw_workbook, properties &C.lxw_doc_properties) Lxw_error
fn C.workbook_set_custom_property_string(workbook &C.lxw_workbook, name &char, value &char) Lxw_error
fn C.workbook_set_custom_property_number(workbook &C.lxw_workbook, name &char, value f64) Lxw_error
fn C.workbook_set_custom_property_integer(workbook &C.lxw_workbook, name &char, value int) Lxw_error
fn C.workbook_set_custom_property_boolean(workbook &C.lxw_workbook, name &char, value u8) Lxw_error
fn C.workbook_set_custom_property_datetime(workbook &C.lxw_workbook, name &char, datetime &C.lxw_datetime) Lxw_error
fn C.workbook_define_name(workbook &C.lxw_workbook, name &char, formula &char) Lxw_error
fn C.workbook_get_default_url_format(workbook &C.lxw_workbook) &C.lxw_format
fn C.workbook_get_worksheet_by_name(workbook &C.lxw_workbook, name &char) &C.lxw_worksheet
fn C.workbook_get_chartsheet_by_name(workbook &C.lxw_workbook, name &char) &C.lxw_chartsheet
fn C.workbook_validate_sheet_name(workbook &C.lxw_workbook, sheetname &char) Lxw_error
fn C.workbook_add_vba_project(workbook &C.lxw_workbook, filename &char) Lxw_error
fn C.workbook_add_signed_vba_project(workbook &C.lxw_workbook, vba_project &char, signature &char) Lxw_error
fn C.workbook_set_vba_name(workbook &C.lxw_workbook, name &char) Lxw_error
fn C.workbook_read_only_recommended(workbook &C.lxw_workbook)
fn C.workbook_unset_default_url_format(workbook &C.lxw_workbook)

// new_workbook create a new workbook object
pub fn new_workbook(filename string) &C.lxw_workbook {
	return C.workbook_new(filename.str)
}

// new_workbook_opt create a new workbook object, and set the workbook options
pub fn new_workbook_opt(filename string, options &WorkBook_Options) &C.lxw_workbook {
	o := &C.lxw_workbook_options{
		constant_memory: options.constant_memory
		tmpdir: options.tmpdir.str
		use_zip64: options.use_zip64
		output_buffer: options.output_buffer.data
		output_buffer_size: &usize(options.output_buffer.len)
	}
	return C.workbook_new_opt(filename.str, o)
}

// add_worksheet add a new worksheet to a workbook
pub fn (b &C.lxw_workbook) add_worksheet(sheetname string) &C.lxw_worksheet {
	return C.workbook_add_worksheet(b, sheetname.str)
}

// add_chartsheet add a new chartsheet to a workbook
pub fn (b &C.lxw_workbook) add_chartsheet(sheetname string) &C.lxw_chartsheet {
	return C.workbook_add_chartsheet(b, sheetname.str)
}

// add_format create a new Format object to formats cells in worksheets
pub fn (b &C.lxw_workbook) add_format() &C.lxw_format {
	return C.workbook_add_format(b)
}

// add_chart create a new chart to be added to a worksheet
pub fn (b &C.lxw_workbook) add_chart(chart_type Lxw_chart_type) &C.lxw_chart {
	return C.workbook_add_chart(b, u8(chart_type))
}

// close close the Workbook object and write the XLSX file
pub fn (b &C.lxw_workbook) close() Lxw_error {
	return C.workbook_close(b)
}

// set_properties set the document properties such as Title, Author etc
pub fn (b &C.lxw_workbook) set_properties(properties &C.lxw_doc_properties) Lxw_error {
	return C.workbook_set_properties(b, properties)
}

// set_custom_property set a custom document property
pub fn (b &C.lxw_workbook) set_custom_property(name string, value All_Data_Type) Lxw_error {
	match value {
		string {
			return C.workbook_set_custom_property_string(b, name.str, value.str)
		}
		f64 {
			return C.workbook_set_custom_property_number(b, name.str, value)
		}
		int {
			return C.workbook_set_custom_property_integer(b, name.str, value)
		}
		bool {
			return C.workbook_set_custom_property_boolean(b, name.str, int(value))
		}
		time.Time {
			datetime := &C.lxw_datetime{
				year: value.year
				month: value.month
				day: value.day
				hour: value.hour
				min: value.minute
				sec: value.second
			}
			return C.workbook_set_custom_property_datetime(b, name.str, datetime)
		}
	}
}

// define_name create a defined name in the workbook to use as a variable
pub fn (b &C.lxw_workbook) define_name(name string, formula string) Lxw_error {
	return C.workbook_define_name(b, name.str, formula.str)
}

// get_default_url_format get the default URL format used with write_url()
pub fn (b &C.lxw_workbook) get_default_url_format() &C.lxw_format {
	return C.workbook_get_default_url_format(b)
}

// get_worksheet_by_name get a worksheet object from its name
pub fn (b &C.lxw_workbook) get_worksheet_by_name(name string) &C.lxw_worksheet {
	return C.workbook_get_worksheet_by_name(b, name.str)
}

// get_chartsheet_by_name get a chartsheet object from its name
pub fn (b &C.lxw_workbook) get_chartsheet_by_name(name string) &C.lxw_chartsheet {
	return C.workbook_get_chartsheet_by_name(b, name.str)
}

// validate_sheet_name validate a worksheet or chartsheet name
pub fn (b &C.lxw_workbook) validate_sheet_name(sheetname string) Lxw_error {
	return C.workbook_validate_sheet_name(b, sheetname.str)
}

// add_vba_project add a vbaProject binary to the Excel workbook
pub fn (b &C.lxw_workbook) add_vba_project(filename string) Lxw_error {
	return C.workbook_add_vba_project(b, filename.str)
}

// add_signed_vba_project add a vbaProject binary and a vbaProjectSignature binary to the Excel workbook
pub fn (b &C.lxw_workbook) add_signed_vba_project(vba_project string, signature string) Lxw_error {
	return C.workbook_add_signed_vba_project(b, vba_project.str, signature.str)
}

// set_vba_name set the VBA name for the workbook
pub fn (b &C.lxw_workbook) set_vba_name(name string) Lxw_error {
	return C.workbook_set_vba_name(b, name.str)
}

// read_only_recommended add a recommendation to open the file in "read-only" mode
pub fn (b &C.lxw_workbook) read_only_recommended() &C.lxw_workbook {
	C.workbook_read_only_recommended(b)
	return unsafe { b }
}

// unset_default_url_format unset the default URL format
pub fn (b &C.lxw_workbook) unset_default_url_format() &C.lxw_workbook {
	C.workbook_unset_default_url_format(b)
	return unsafe { b }
}
