module vxlsx

// ===========chartsheet functions===========
fn C.chartsheet_set_chart(chartsheet &C.lxw_chartsheet, chart &C.lxw_chart) Lxw_error
fn C.chartsheet_activate(chartsheet &C.lxw_chartsheet)
fn C.chartsheet_select(chartsheet &C.lxw_chartsheet)
fn C.chartsheet_hide(chartsheet &C.lxw_chartsheet)
fn C.chartsheet_set_first_sheet(chartsheet &C.lxw_chartsheet)
fn C.chartsheet_set_tab_color(chartsheet &C.lxw_chartsheet, color u32)
fn C.chartsheet_protect(chartsheet &C.lxw_chartsheet, password &char, options &C.lxw_protection)
fn C.chartsheet_set_zoom(chartsheet &C.lxw_chartsheet, scale u16)
fn C.chartsheet_set_landscape(chartsheet &C.lxw_chartsheet)
fn C.chartsheet_set_portrait(chartsheet &C.lxw_chartsheet)
fn C.chartsheet_set_paper(chartsheet &C.lxw_chartsheet, paper_type u8)
fn C.chartsheet_set_margins(chartsheet &C.lxw_chartsheet, left f64, right f64, top f64, bottom f64)
fn C.chartsheet_set_header(chartsheet &C.lxw_chartsheet, string_ &char) Lxw_error
fn C.chartsheet_set_footer(chartsheet &C.lxw_chartsheet, string_ &char) Lxw_error
fn C.chartsheet_set_header_opt(chartsheet &C.lxw_chartsheet, string_ &char, options &C.lxw_header_footer_options) Lxw_error
fn C.chartsheet_set_footer_opt(chartsheet &C.lxw_chartsheet, string_ &char, options &C.lxw_header_footer_options) Lxw_error

// set_chart insert a chart object into a chartsheet
pub fn (c &C.lxw_chartsheet) set_chart(chart &C.lxw_chart) Lxw_error {
	return C.chartsheet_set_chart(c, chart)
}

// activate make a chartsheet the active, i.e., visible chartsheet
pub fn (c &C.lxw_chartsheet) activate() &C.lxw_chartsheet {
	C.chartsheet_activate(c)
	return unsafe { c }
}

// @select set a chartsheet tab as selected
pub fn (c &C.lxw_chartsheet) @select() &C.lxw_chartsheet {
	C.chartsheet_select(c)
	return unsafe { c }
}

// hide hide the current chartsheet
pub fn (c &C.lxw_chartsheet) hide() &C.lxw_chartsheet {
	C.chartsheet_hide(c)
	return unsafe { c }
}

// set_first_sheet set current chartsheet as the first visible sheet tab
pub fn (c &C.lxw_chartsheet) set_first_sheet() &C.lxw_chartsheet {
	C.chartsheet_set_first_sheet(c)
	return unsafe { c }
}

// set_tab_color set the color of the chartsheet tab
pub fn (c &C.lxw_chartsheet) set_tab_color(color u32) &C.lxw_chartsheet {
	C.chartsheet_set_tab_color(c, color)
	return unsafe { c }
}

// protect protect elements of a chartsheet from modification
pub fn (c &C.lxw_chartsheet) protect(password string, options &C.lxw_protection) &C.lxw_chartsheet {
	C.chartsheet_protect(c, password.str, options)
	return unsafe { c }
}

// set_zoom set the chartsheet zoom factor, `scale`(10 to 400)
pub fn (c &C.lxw_chartsheet) set_zoom(scale u16) &C.lxw_chartsheet {
	C.chartsheet_set_zoom(c, scale)
	return unsafe { c }
}

// set_landscape set the page orientation as landscape
pub fn (c &C.lxw_chartsheet) set_landscape() &C.lxw_chartsheet {
	C.chartsheet_set_landscape(c)
	return unsafe { c }
}

// set_portrait set the page orientation as portrait
pub fn (c &C.lxw_chartsheet) set_portrait() &C.lxw_chartsheet {
	C.chartsheet_set_portrait(c)
	return unsafe { c }
}

// set_paper set the paper type for printing
pub fn (c &C.lxw_chartsheet) set_paper(paper_type Paper_format_type) &C.lxw_chartsheet {
	C.chartsheet_set_paper(c, u8(paper_type))
	return unsafe { c }
}

// set_margins set the chartsheet margins for the printed page
pub fn (c &C.lxw_chartsheet) set_margins(left f64, right f64, top f64, bottom f64) &C.lxw_chartsheet {
	C.chartsheet_set_margins(c, left, right, top, bottom)
	return unsafe { c }
}

// set_header set the printed page header caption
pub fn (c &C.lxw_chartsheet) set_header(str string) Lxw_error {
	return C.chartsheet_set_header(c, str.str)
}

// set_footer set the printed page footer caption
pub fn (c &C.lxw_chartsheet) set_footer(str string) Lxw_error {
	return C.chartsheet_set_footer(c, str.str)
}

// set_header_opt set the printed page header caption with additional options
pub fn (c &C.lxw_chartsheet) set_header_opt(str string, options &C.lxw_header_footer_options) Lxw_error {
	return C.chartsheet_set_header_opt(c, str.str, options)
}

// set_footer_opt set the printed page footer caption with additional options
pub fn (c &C.lxw_chartsheet) set_footer_opt(str string, options &C.lxw_header_footer_options) Lxw_error {
	return C.chartsheet_set_footer_opt(c, str.str, options)
}
