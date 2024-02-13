module vxlsx

import time

@[typedef]
struct C.lxw_datetime {
	year  int
	month int
	day   int
	hour  int
	min   int
	sec   f64
}

@[typedef]
struct C.lxw_row {}

@[typedef]
struct C.lxw_cell {}

@[typedef]
struct C.lxw_worksheet {}

@[typedef]
struct C.lxw_chart {}

@[typedef]
struct C.lxw_chartsheet {}

@[typedef]
pub struct C.lxw_protection {
pub mut:
	no_select_locked_cells   u8
	no_select_unlocked_cells u8
	format_cells             u8
	format_columns           u8
	format_rows              u8
	insert_columns           u8
	insert_rows              u8
	insert_hyperlinks        u8
	delete_columns           u8
	delete_rows              u8
	sort                     u8
	autofilter               u8
	pivot_tables             u8
	scenarios                u8
	objects                  u8
	no_content               u8
	no_objects               u8
}

@[typedef]
pub struct C.lxw_header_footer_options {
pub mut:
	margin       f64
	image_left   &char
	image_center &char
	image_right  &char
}

@[typedef]
pub struct C.lxw_table_options {
pub mut:
	name              &char
	no_header_row     u8
	no_autofilter     u8
	no_banded_rows    u8
	banded_columns    u8
	first_column      u8
	last_column       u8
	style_type        u8
	style_type_number u8
	total_row         u8
	columns           voidptr
}

@[typedef]
pub struct C.lxw_button_options {
pub mut:
	caption     &char
	macro       &char
	description &char
	width       u16
	height      u16
	x_scale     f64
	y_scale     f64
	x_offset    int
	y_offset    int
}

@[typedef]
struct C.lxw_conditional_format {}

@[typedef]
struct C.lxw_data_validation {}

@[typedef]
struct C.lxw_filter_rule {}

@[typedef]
pub struct C.lxw_chart_options {
pub mut:
	x_offset        int
	y_offset        int
	x_scale         f64
	y_scale         f64
	object_position u8
	description     &char
	decorative      u8
}

@[typedef]
pub struct C.lxw_image_options {
pub mut:
	x_offset        int
	y_offset        int
	x_scale         f64
	y_scale         f64
	object_position u8
	description     &char
	decorative      u8
	url             &char
	tip             &char
}

@[typedef]
pub struct C.lxw_row_col_options {
pub mut:
	hidden    u8
	level     u8
	collapsed u8
}

@[typedef]
pub struct C.lxw_comment_options {
pub mut:
	visible     u8
	author      &char
	width       u16
	height      u16
	x_scale     f64
	y_scale     f64
	color       u32
	font_name   &char
	font_size   f64
	font_family u8
	start_row   u32
	start_col   u16
	x_offset    int
	y_offset    int
}

@[typedef]
struct C.lxw_rich_string_tuple {}

// Lxw_gridlines gridline options using in worksheet.gridlines()
pub enum Lxw_gridlines {
	hide_all_gridlines    = 0 // Hide screen and print gridlines
	show_screen_gridlines // Show screen gridlines
	show_print_gridlines // Show print gridlines
	show_all_gridlines // Show screen and print gridlines
}

// Lxw_validation_boolean data validation property values
pub enum Lxw_validation_boolean {
	validation_default
	validation_off // Turn a data validation property off
	validation_on // Turn a data validation property on. Data validation properties are generally on by default
}

// Lxw_validation_types data validation types
pub enum Lxw_validation_types {
	validation_type_none
	validation_type_integer // Restrict cell input to whole/integer numbers only
	validation_type_integer_formula // Restrict cell input to whole/integer numbers only, using a cell reference
	validation_type_decimal // Restrict cell input to decimal numbers only
	validation_type_decimal_formula // Restrict cell input to decimal numbers only, using a cell reference
	validation_type_list // Restrict cell input to a list of strings in a dropdown
	validation_type_list_formula // Restrict cell input to a list of strings in a dropdown, using a cell range
	validation_type_date // Restrict cell input to date values only, using a lxw_datetime type
	validation_type_date_formula // Restrict cell input to date values only, using a cell reference
	validation_type_date_number // Restrict cell input to date values only, as a serial number
	validation_type_time // Restrict cell input to time values only, using a lxw_datetime type
	validation_type_time_formula // Restrict cell input to time values only, using a cell reference
	validation_type_time_number // Restrict cell input to time values only, as a serial number
	validation_type_length // Restrict cell input to strings of defined length, using a cell reference
	validation_type_length_formula // Restrict cell input to strings of defined length, using a cell reference
	validation_type_custom_formula // Restrict cell to input controlled by a custom formula that returns `TRUE/FALSE`
	validation_type_any // Allow any type of input. Mainly only useful for pop-up messages
}

// Lxw_validation_criteria data validation criteria uses to control the selection of data
pub enum Lxw_validation_criteria {
	validation_criteria_none
	validation_criteria_between // Select data between two values
	validation_criteria_not_between // Select data that is not between two values
	validation_criteria_equal_to // Select data equal to a value
	validation_criteria_not_equal_to // Select data not equal to a value
	validation_criteria_greater_than // Select data greater than a value
	validation_criteria_less_than // Select data less than a value
	validation_criteria_greater_than_or_equal_to // Select data greater than or equal to a value
	validation_criteria_less_than_or_equal_to // Select data less than or equal to a value
}

// Lxw_validation_error_types data validation error types for pop-up messages
pub enum Lxw_validation_error_types {
	validation_error_type_stop // Show a "Stop" data validation pop-up message. This is the default
	validation_error_type_warning // Show an "Error" data validation pop-up message
	validation_error_type_information // Show an "Information" data validation pop-up message
}

// Lxw_comment_display_types set the display type for a cell comment. This is hidden by default but can be set to visible with the worksheet.show_comments() function
pub enum Lxw_comment_display_types {
	lxw_comment_display_default // Default to the worksheet default which can be hidden or visible
	lxw_comment_display_hidden // Hide the cell comment. Usually the default
	lxw_comment_display_visible // Show the cell comment. Can also be set for the worksheet with the worksheet.show_comments() function
}

// Lxw_conditional_format_types type definitions for conditional formats
pub enum Lxw_conditional_format_types {
	conditional_type_none
	conditional_type_cell // The Cell type is the most common conditional formatting type. It is used when a format is applied to a cell based on a simple criterion
	conditional_type_text // The Text type is used to specify Excel's "Specific Text" style conditional format
	conditional_type_time_period // The Time Period type is used to specify Excel's "Dates Occurring" style conditional format
	conditional_type_average // The Average type is used to specify Excel's "Average" style conditional format
	conditional_type_duplicate // The Duplicate type is used to highlight duplicate cells in a range
	conditional_type_unique // The Unique type is used to highlight unique cells in a range
	conditional_type_top // The Top type is used to specify the top n values by number or percentage in a range
	conditional_type_bottom // The Bottom type is used to specify the bottom n values by number or percentage in a range
	conditional_type_blanks // The Blanks type is used to highlight blank cells in a range
	conditional_type_no_blanks // The No Blanks type is used to highlight non blank cells in a range
	conditional_type_errors // The Errors type is used to highlight error cells in a range
	conditional_type_no_errors // The No Errors type is used to highlight non error cells in a range
	conditional_type_formula // The Formula type is used to specify a conditional format based on a user defined formula
	conditional_2_color_scale // The 2 Color Scale type is used to specify Excel's "2 Color Scale" style conditional format
	conditional_3_color_scale // The 3 Color Scale type is used to specify Excel's "3 Color Scale" style conditional format
	conditional_data_bar // The Data Bar type is used to specify Excel's "Data Bar" style conditional format
	conditional_type_icon_sets // The Icon Set type is used to specify a conditional format with a set of icons such as traffic lights or arrows
	conditional_type_last
}

// Lxw_conditional_criteria the criteria used in a conditional format
pub enum Lxw_conditional_criteria {
	conditional_criteria_none
	conditional_criteria_equal_to // Format cells equal to a value
	conditional_criteria_not_equal_to // Format cells not equal to a value
	conditional_criteria_greater_than // Format cells greater than a value
	conditional_criteria_less_than // Format cells less than a value
	conditional_criteria_greater_than_or_equal_to // Format cells greater than or equal to a value
	conditional_criteria_less_than_or_equal_to // Format cells less than or equal to a value
	conditional_criteria_between // Format cells between two values
	conditional_criteria_not_between // Format cells that is not between two values
	conditional_criteria_text_containing // Format cells that contain the specified text
	conditional_criteria_text_not_containing // Format cells that don't contain the specified text
	conditional_criteria_text_begins_with // Format cells that begin with the specified text
	conditional_criteria_text_ends_with // Format cells that end with the specified text
	conditional_criteria_time_period_yesterday // Format cells with a date of yesterday
	conditional_criteria_time_period_today // Format cells with a date of today
	conditional_criteria_time_period_tomorrow // Format cells with a date of tomorrow
	conditional_criteria_time_period_last_7_days // Format cells with a date in the last 7 days
	conditional_criteria_time_period_last_week // Format cells with a date in the last week
	conditional_criteria_time_period_this_week // Format cells with a date in the current week
	conditional_criteria_time_period_next_week // Format cells with a date in the next week
	conditional_criteria_time_period_last_month // Format cells with a date in the last month
	conditional_criteria_time_period_this_month // Format cells with a date in the current month
	conditional_criteria_time_period_next_month // Format cells with a date in the next month
	conditional_criteria_average_above // Format cells above the average for the range
	conditional_criteria_average_below // Format cells below the average for the range
	conditional_criteria_average_above_or_equal // Format cells above or equal to the average for the range
	conditional_criteria_average_below_or_equal // Format cells below or equal to the average for the range
	conditional_criteria_average_1_std_dev_above // Format cells 1 standard deviation above the average for the range
	conditional_criteria_average_1_std_dev_below // Format cells 1 standard deviation below the average for the range
	conditional_criteria_average_2_std_dev_above // Format cells 2 standard deviation above the average for the range
	conditional_criteria_average_2_std_dev_below // Format cells 2 standard deviation below the average for the range
	conditional_criteria_average_3_std_dev_above // Format cells 3 standard deviation above the average for the range
	conditional_criteria_average_3_std_dev_below // Format cells 3 standard deviation below the average for the range
	conditional_criteria_top_or_bottom_percent // Format cells in the top of bottom percentage
}

// Lxw_conditional_format_rule_types conditional format rule types
pub enum Lxw_conditional_format_rule_types {
	conditional_rule_type_none
	conditional_rule_type_minimum // Conditional format rule type: matches the minimum values in the range. Can only be applied to min_rule_type
	conditional_rule_type_number // Conditional format rule type: use a number to set the bound
	conditional_rule_type_percent // Conditional format rule type: use a percentage to set the bound
	conditional_rule_type_percentile // Conditional format rule type: use a percentile to set the bound
	conditional_rule_type_formula // Conditional format rule type: use a formula to set the bound
	conditional_rule_type_maximum // Conditional format rule type: matches the maximum values in the range. Can only be applied to max_rule_type
	conditional_rule_type_auto_min // Used internally for Excel2010 bars. Not documented
	conditional_rule_type_auto_max // Used internally for Excel2010 bars. Not documented
}

// Lxw_conditional_format_bar_direction conditional format data bar directions
pub enum Lxw_conditional_format_bar_direction {
	conditional_bar_direction_context // Data bar direction is set by Excel based on the context of the data displayed
	conditional_bar_direction_right_to_left // Data bar direction is from right to left
	conditional_bar_direction_left_to_right // Data bar direction is from left to right
}

// Lxw_conditional_bar_axis_position conditional format data bar axis options
pub enum Lxw_conditional_bar_axis_position {
	conditional_bar_axis_automatic // Data bar axis position is set by Excel based on the context of the data displayed
	conditional_bar_axis_midpoint // Data bar axis position is set at the midpoint
	conditional_bar_axis_none // Data bar axis is turned off
}

// Lxw_conditional_icon_types icon types used in the #lxw_conditional_format icon_style field
pub enum Lxw_conditional_icon_types {
	conditional_icons_3_arrows_colored // Icon style: 3 colored arrows showing up, sideways and down
	conditional_icons_3_arrows_gray // Icon style: 3 gray arrows showing up, sideways and down
	conditional_icons_3_flags // Icon style: 3 colored flags in red, yellow and green
	conditional_icons_3_traffic_lights_unrimmed // Icon style: 3 traffic lights - rounded
	conditional_icons_3_traffic_lights_rimmed // Icon style: 3 traffic lights with a rim - squarish
	conditional_icons_3_signs // Icon style: 3 colored shapes - a circle, triangle and diamond
	conditional_icons_3_symbols_circled // Icon style: 3 circled symbols with tick mark, exclamation and cross
	conditional_icons_3_symbols_uncircled // Icon style: 3 symbols with tick mark, exclamation and cross
	conditional_icons_4_arrows_colored // Icon style: 4 colored arrows showing up, diagonal up, diagonal down and down
	conditional_icons_4_arrows_gray // Icon style: 4 gray arrows showing up, diagonal up, diagonal down and down
	conditional_icons_4_red_to_black // Icon style: 4 circles in 4 colors going from red to black
	conditional_icons_4_ratings // Icon style: 4 histogram ratings
	conditional_icons_4_traffic_lights // Icon style: 4 traffic lights
	conditional_icons_5_arrows_colored // Icon style: 5 colored arrows showing up, diagonal up, sideways, diagonal down and down
	conditional_icons_5_arrows_gray // Icon style: 5 gray arrows showing up, diagonal up, sideways, diagonal down and down
	conditional_icons_5_ratings // Icon style: 5 histogram ratings
	conditional_icons_5_quarters // Icon style: 5 quarters, from 0 to 4 quadrants filled
}

// Lxw_table_style_type the type of table style
pub enum Lxw_table_style_type {
	table_style_type_default
	table_style_type_light // Light table style
	table_style_type_medium // Medium table style
	table_style_type_dark // Dark table style
}

// Lxw_table_total_functions standard Excel functions for totals in tables
pub enum Lxw_table_total_functions {
	table_function_none       = 0
	table_function_average    = 101 // Use the average function as the table total
	table_function_count_nums = 102 // Use the count numbers function as the table total
	table_function_count      = 103 // Use the count function as the table total
	table_function_max        = 104 // Use the max function as the table total
	table_function_min        = 105 // Use the min function as the table total
	table_function_std_dev    = 107 // Use the standard deviation function as the table total
	table_function_sum        = 109 // Use the sum function as the table total
	table_function_var        = 110 // Use the var function as the table total
}

// Lxw_filter_criteria the criteria used in autofilter rules
pub enum Lxw_filter_criteria {
	filter_criteria_none
	filter_criteria_equal_to // Filter cells equal to a value
	filter_criteria_not_equal_to // Filter cells not equal to a value
	filter_criteria_greater_than // Filter cells greater than a value
	filter_criteria_less_than // Filter cells less than a value
	filter_criteria_greater_than_or_equal_to // Filter cells greater than or equal to a value
	filter_criteria_less_than_or_equal_to // Filter cells less than or equal to a value
	filter_criteria_blanks // Filter cells that are blank
	filter_criteria_non_blanks // Filter cells that are not blank
}

// Lxw_filter_operator and/or operator when using 2 filter rules
pub enum Lxw_filter_operator {
	filter_and // Logical "and" of 2 filter rules
	filter_or // Logical "or" of 2 filter rules
}

// Lxw_filter_type internal filter types
pub enum Lxw_filter_type {
	filter_type_none
	filter_type_single
	filter_type_and
	filter_type_or
	filter_type_string_list
}

// Lxw_object_position options to control the positioning of worksheet objects such as images or charts
pub enum Lxw_object_position {
	object_position_default // Default positioning for the object
	object_move_and_size // Move and size the worksheet object with the cells
	object_move_dont_size // Move but don't size the worksheet object with the cells
	object_dont_move_dont_size // Don't move or size the worksheet object with the cells
	object_move_and_size_after // Same as #LXW_OBJECT_MOVE_AND_SIZE except libxlsxwriter applies hidden cells after the object is inserted
}

// Lxw_ignore_errors options for ignoring worksheet errors/warnings. See worksheet.ignore_errors()
pub enum Lxw_ignore_errors {
	ignore_number_stored_as_text = 1 // Turn off errors/warnings for numbers stores as text
	ignore_eval_error // Turn off errors/warnings for formula errors (such as divide by zero)
	ignore_formula_differs // Turn off errors/warnings for formulas that differ from surrounding formulas
	ignore_formula_range // Turn off errors/warnings for formulas that omit cells in a range
	ignore_formula_unlocked // Turn off errors/warnings for unlocked cells that contain formulas
	ignore_empty_cell_reference // Turn off errors/warnings for formulas that refer to empty cells
	ignore_list_data_validation // Turn off errors/warnings for cells in a table that do not comply with applicable data validation rules
	ignore_calculated_column // Turn off errors/warnings for cell formulas that differ from the column formula
	ignore_two_digit_text_year // Turn off errors/warnings for formulas that contain a two digit text representation of a year
	ignore_last_option
}

pub enum Cell_types {
	number_cell                = 1
	string_cell
	inline_string_cell
	inline_rich_string_cell
	formula_cell
	array_formula_cell
	dynamic_array_formula_cell
	blank_cell
	boolean_cell
	comment
	hyperlink_url
	hyperlink_internal
	hyperlink_external
}

pub enum Pane_types {
	no_panes           = 0
	freeze_panes
	split_panes
	freeze_split_panes
}

pub enum Lxw_image_position {
	header_left   = 0
	header_center
	header_right
	footer_left
	footer_center
	footer_right
}

pub enum Paper_format_type {
	paper_format_type_printer_default // Printer default
	paper_format_type_letter // 8 1/2 x 11 in
	paper_format_type_letter_small // 8 1/2 x 11 in
	paper_format_type_tabloid // 11 x 17 in
	paper_format_type_ledger // 17 x 11 in
	paper_format_type_legal // 8 1/2 x 14 in
	paper_format_type_statement // 5 1/2 x 8 1/2 in
	paper_format_type_executive // 7 1/4 x 10 1/2 in
	paper_format_type_a3 // 297 x 420 mm
	paper_format_type_a4 // 210 x 297 mm
	paper_format_type_a4_small // 210 x 297 mm
	paper_format_type_a5 // 148 x 210 mm
	paper_format_type_b4 // 250 x 354 mm
	paper_format_type_b5 // 182 x 257 mm
	paper_format_type_folio // 8 1/2 x 13 in
	paper_format_type_quarto // 215 x 275 mm
	paper_format_type_10x14_in // 10x14 in
	paper_format_type_11x17_in // 11x17 in
	paper_format_type_note // 8 1/2 x 11 in
	paper_format_type_envelope_9 // 3 7/8 x 8 7/8
	paper_format_type_envelope_10 // 4 1/8 x 9 1/2
	paper_format_type_envelope_11 // 4 1/2 x 10 3/8
	paper_format_type_envelope_12 // 4 3/4 x 11
	paper_format_type_envelope_14 // 5 x 11 1/2
	paper_format_type_c_size_sheet // —
	paper_format_type_d_size_sheet // —
	paper_format_type_e_size_sheet // —
	paper_format_type_envelope_dl // 110 x 220 mm
	paper_format_type_envelope_c3 // 324 x 458 mm
	paper_format_type_envelope_c4 // 229 x 324 mm
	paper_format_type_envelope_c5 // 162 x 229 mm
	paper_format_type_envelope_c6 // 114 x 162 mm
	paper_format_type_envelope_c65 // 114 x 229 mm
	paper_format_type_envelope_b4 // 250 x 353 mm
	paper_format_type_envelope_b5 // 176 x 250 mm
	paper_format_type_envelope_b6 // 176 x 125 mm
	paper_format_type_envelope_110 // 110 x 230 mm
	paper_format_type_monarch // 3.875 x 7.5 in
	paper_format_type_envelope // 3 5/8 x 6 1/2 in
	paper_format_type_fanfold // 14 7/8 x 11 in
	paper_format_type_german_std_fanfold // 8 1/2 x 12 in
	paper_format_type_german_legal_fanfold // 8 1/2 x 13 in
}

// ===========worksheet functions===========
fn C.worksheet_write_number(worksheet &C.lxw_worksheet, row u32, col u16, number f64, format &C.lxw_format) Lxw_error
fn C.worksheet_write_string(worksheet &C.lxw_worksheet, row u32, col u16, string_ &char, format &C.lxw_format) Lxw_error
fn C.worksheet_write_formula(worksheet &C.lxw_worksheet, row u32, col u16, formula &char, format &C.lxw_format) Lxw_error
fn C.worksheet_write_array_formula(worksheet &C.lxw_worksheet, first_row u32, first_col u16, last_row u32, last_col u16, formula &char, format &C.lxw_format) Lxw_error
fn C.worksheet_write_dynamic_array_formula(worksheet &C.lxw_worksheet, first_row u32, first_col u16, last_row u32, last_col u16, formula &char, format &C.lxw_format) Lxw_error
fn C.worksheet_write_dynamic_formula(worksheet &C.lxw_worksheet, row u32, col u16, formula &char, format &C.lxw_format) Lxw_error
fn C.worksheet_write_array_formula_num(worksheet &C.lxw_worksheet, first_row u32, first_col u16, last_row u32, last_col u16, formula &char, format &C.lxw_format, result f64) Lxw_error
fn C.worksheet_write_dynamic_array_formula_num(worksheet &C.lxw_worksheet, first_row u32, first_col u16, last_row u32, last_col u16, formula &char, format &C.lxw_format, result f64) Lxw_error
fn C.worksheet_write_dynamic_formula_num(worksheet &C.lxw_worksheet, row u32, col u16, formula &char, format &C.lxw_format, result f64) Lxw_error
fn C.worksheet_write_datetime(worksheet &C.lxw_worksheet, row u32, col u16, datetime &C.lxw_datetime, format &C.lxw_format) Lxw_error
fn C.worksheet_write_unixtime(worksheet &C.lxw_worksheet, row u32, col u16, unixtime i64, format &C.lxw_format) Lxw_error
fn C.worksheet_write_url(worksheet &C.lxw_worksheet, row u32, col u16, url &char, format &C.lxw_format) Lxw_error
fn C.worksheet_write_url_opt(worksheet &C.lxw_worksheet, row_num u32, col_num u16, url &char, format &C.lxw_format, string_ &char, tooltip &char) Lxw_error
fn C.worksheet_write_boolean(worksheet &C.lxw_worksheet, row u32, col u16, value int, format &C.lxw_format) Lxw_error
fn C.worksheet_write_blank(worksheet &C.lxw_worksheet, row u32, col u16, format &C.lxw_format) Lxw_error
fn C.worksheet_write_formula_num(worksheet &C.lxw_worksheet, row u32, col u16, formula &char, format &C.lxw_format, result f64) Lxw_error
fn C.worksheet_write_formula_str(worksheet &C.lxw_worksheet, row u32, col u16, formula &char, format &C.lxw_format, result &char) Lxw_error
fn C.worksheet_write_rich_string(worksheet &C.lxw_worksheet, row u32, col u16, rich_string &&C.lxw_rich_string_tuple, format &C.lxw_format) Lxw_error
fn C.worksheet_write_comment(worksheet &C.lxw_worksheet, row u32, col u16, string_ &char) Lxw_error
fn C.worksheet_write_comment_opt(worksheet &C.lxw_worksheet, row u32, col u16, string_ &char, options &C.lxw_comment_options) Lxw_error
fn C.worksheet_set_row(worksheet &C.lxw_worksheet, row u32, height f64, format &C.lxw_format) Lxw_error
fn C.worksheet_set_row_opt(worksheet &C.lxw_worksheet, row u32, height f64, format &C.lxw_format, options &C.lxw_row_col_options) Lxw_error
fn C.worksheet_set_row_pixels(worksheet &C.lxw_worksheet, row u32, pixels u32, format &C.lxw_format) Lxw_error
fn C.worksheet_set_row_pixels_opt(worksheet &C.lxw_worksheet, row u32, pixels u32, format &C.lxw_format, options &C.lxw_row_col_options) Lxw_error
fn C.worksheet_set_column(worksheet &C.lxw_worksheet, first_col u16, last_col u16, width f64, format &C.lxw_format) Lxw_error
fn C.worksheet_set_column_opt(worksheet &C.lxw_worksheet, first_col u16, last_col u16, width f64, format &C.lxw_format, options &C.lxw_row_col_options) Lxw_error
fn C.worksheet_set_column_pixels(worksheet &C.lxw_worksheet, first_col u16, last_col u16, pixels u32, format &C.lxw_format) Lxw_error
fn C.worksheet_set_column_pixels_opt(worksheet &C.lxw_worksheet, first_col u16, last_col u16, pixels u32, format &C.lxw_format, options &C.lxw_row_col_options) Lxw_error
fn C.worksheet_insert_image(worksheet &C.lxw_worksheet, row u32, col u16, filename &char) Lxw_error
fn C.worksheet_insert_image_opt(worksheet &C.lxw_worksheet, row u32, col u16, filename &char, options &C.lxw_image_options) Lxw_error
fn C.worksheet_insert_image_buffer(worksheet &C.lxw_worksheet, row u32, col u16, image_buffer &u8, image_size usize) Lxw_error
fn C.worksheet_insert_image_buffer_opt(worksheet &C.lxw_worksheet, row u32, col u16, image_buffer &u8, image_size usize, options &C.lxw_image_options) Lxw_error
fn C.worksheet_set_background(worksheet &C.lxw_worksheet, filename &char) Lxw_error
fn C.worksheet_set_background_buffer(worksheet &C.lxw_worksheet, image_buffer &u8, image_size usize) Lxw_error
fn C.worksheet_insert_chart(worksheet &C.lxw_worksheet, row u32, col u16, chart &C.lxw_chart) Lxw_error
fn C.worksheet_insert_chart_opt(worksheet &C.lxw_worksheet, row u32, col u16, chart &C.lxw_chart, user_options &C.lxw_chart_options) Lxw_error
fn C.worksheet_merge_range(worksheet &C.lxw_worksheet, first_row u32, first_col u16, last_row u32, last_col u16, string_ &char, format &C.lxw_format) Lxw_error
fn C.worksheet_autofilter(worksheet &C.lxw_worksheet, first_row u32, first_col u16, last_row u32, last_col u16) Lxw_error
fn C.worksheet_filter_column(worksheet &C.lxw_worksheet, col u16, rule &C.lxw_filter_rule) Lxw_error
fn C.worksheet_filter_column2(worksheet &C.lxw_worksheet, col u16, rule1 &C.lxw_filter_rule, rule2 &C.lxw_filter_rule, and_or u8) Lxw_error
fn C.worksheet_filter_list(worksheet &C.lxw_worksheet, col u16, list &&u8) Lxw_error
fn C.worksheet_data_validation_cell(worksheet &C.lxw_worksheet, row u32, col u16, validation &C.lxw_data_validation) Lxw_error
fn C.worksheet_data_validation_range(worksheet &C.lxw_worksheet, first_row u32, first_col u16, last_row u32, last_col u16, validation &C.lxw_data_validation) Lxw_error
fn C.worksheet_conditional_format_cell(worksheet &C.lxw_worksheet, row u32, col u16, conditional_format &C.lxw_conditional_format) Lxw_error
fn C.worksheet_conditional_format_range(worksheet &C.lxw_worksheet, first_row u32, first_col u16, last_row u32, last_col u16, conditional_format &C.lxw_conditional_format) Lxw_error
fn C.worksheet_insert_button(worksheet &C.lxw_worksheet, row u32, col u16, options &C.lxw_button_options) Lxw_error
fn C.worksheet_add_table(worksheet &C.lxw_worksheet, first_row u32, first_col u16, last_row u32, last_col u16, options &C.lxw_table_options) Lxw_error
fn C.worksheet_activate(worksheet &C.lxw_worksheet)
fn C.worksheet_select(worksheet &C.lxw_worksheet)
fn C.worksheet_hide(worksheet &C.lxw_worksheet)
fn C.worksheet_set_first_sheet(worksheet &C.lxw_worksheet)
fn C.worksheet_freeze_panes(worksheet &C.lxw_worksheet, row u32, col u16)
fn C.worksheet_split_panes(worksheet &C.lxw_worksheet, vertical f64, horizontal f64)
fn C.worksheet_freeze_panes_opt(worksheet &C.lxw_worksheet, first_row u32, first_col u16, top_row u32, left_col u16, type_ u8)
fn C.worksheet_split_panes_opt(worksheet &C.lxw_worksheet, vertical f64, horizontal f64, top_row u32, left_col u16)
fn C.worksheet_set_selection(worksheet &C.lxw_worksheet, first_row u32, first_col u16, last_row u32, last_col u16)
fn C.worksheet_set_top_left_cell(worksheet &C.lxw_worksheet, row u32, col u16)
fn C.worksheet_set_landscape(worksheet &C.lxw_worksheet)
fn C.worksheet_set_portrait(worksheet &C.lxw_worksheet)
fn C.worksheet_set_page_view(worksheet &C.lxw_worksheet)
fn C.worksheet_set_paper(worksheet &C.lxw_worksheet, paper_type u8)
fn C.worksheet_set_margins(worksheet &C.lxw_worksheet, left f64, right f64, top f64, bottom f64)
fn C.worksheet_set_header(worksheet &C.lxw_worksheet, string_ &char) Lxw_error
fn C.worksheet_set_footer(worksheet &C.lxw_worksheet, string_ &char) Lxw_error
fn C.worksheet_set_header_opt(worksheet &C.lxw_worksheet, string_ &char, options &C.lxw_header_footer_options) Lxw_error
fn C.worksheet_set_footer_opt(worksheet &C.lxw_worksheet, string_ &char, options &C.lxw_header_footer_options) Lxw_error
fn C.worksheet_set_h_pagebreaks(worksheet &C.lxw_worksheet, breaks &u32) Lxw_error
fn C.worksheet_set_v_pagebreaks(worksheet &C.lxw_worksheet, breaks &u16) Lxw_error
fn C.worksheet_print_across(worksheet &C.lxw_worksheet)
fn C.worksheet_set_zoom(worksheet &C.lxw_worksheet, scale u16)
fn C.worksheet_gridlines(worksheet &C.lxw_worksheet, option u8)
fn C.worksheet_center_horizontally(worksheet &C.lxw_worksheet)
fn C.worksheet_center_vertically(worksheet &C.lxw_worksheet)
fn C.worksheet_print_row_col_headers(worksheet &C.lxw_worksheet)
fn C.worksheet_repeat_rows(worksheet &C.lxw_worksheet, first_row u32, last_row u32) Lxw_error
fn C.worksheet_repeat_columns(worksheet &C.lxw_worksheet, first_col u16, last_col u16) Lxw_error
fn C.worksheet_print_area(worksheet &C.lxw_worksheet, first_row u32, first_col u16, last_row u32, last_col u16) Lxw_error
fn C.worksheet_fit_to_pages(worksheet &C.lxw_worksheet, width u16, height u16)
fn C.worksheet_set_start_page(worksheet &C.lxw_worksheet, start_page u16)
fn C.worksheet_set_print_scale(worksheet &C.lxw_worksheet, scale u16)
fn C.worksheet_print_black_and_white(worksheet &C.lxw_worksheet)
fn C.worksheet_right_to_left(worksheet &C.lxw_worksheet)
fn C.worksheet_hide_zero(worksheet &C.lxw_worksheet)
fn C.worksheet_set_tab_color(worksheet &C.lxw_worksheet, color u32)
fn C.worksheet_protect(worksheet &C.lxw_worksheet, password &char, options &C.lxw_protection)
fn C.worksheet_outline_settings(worksheet &C.lxw_worksheet, visible u8, symbols_below u8, symbols_right u8, auto_style u8)
fn C.worksheet_set_default_row(worksheet &C.lxw_worksheet, height f64, hide_unused_rows u8)
fn C.worksheet_set_vba_name(worksheet &C.lxw_worksheet, name &char) Lxw_error
fn C.worksheet_show_comments(worksheet &C.lxw_worksheet)
fn C.worksheet_set_comments_author(worksheet &C.lxw_worksheet, author &char)
fn C.worksheet_ignore_errors(worksheet &C.lxw_worksheet, type_ u8, range &char) Lxw_error

// write_struct write a struct array to xlsx file
// Example 1:
// struct Expense {
// 	item string              @[row:0;title:'ITEM'] // one record per row,from row 0; title is 'ITEM'
// 	cost vxlsx.All_Data_Type @[title:'Cost']       // title is 'Cost'
// }
// Example 2:
// struct Expense {
// 	item string              @[col:0;title:'ITEM'] // one record per col,from col 0; title is 'ITEM'
// 	cost vxlsx.All_Data_Type                       // title is 'cost'
// }
pub fn (s &C.lxw_worksheet) write_struct[T](data []T) {
	mut row := u32(0)
	mut col := u16(0)
	mut row_inc := false
	mut write_title := false
	$for field in T.fields {
		for a in field.attrs {
			if a.starts_with('col:') {
				f := a.split_any(':')
				col = f[1].trim_space().u16()
				row_inc = true
			} else if a.starts_with('row:') {
				f := a.split_any(':')
				row = f[1].trim_space().u16()
				row_inc = false
			}
			if a.starts_with('title') {
				write_title = true
			}
		}
	}
	mut row_reset := row
	mut col_reset := col
	mut has_title := false
	$for field in T.fields {
		if write_title {
			has_title = false
			for a in field.attrs {
				if a.starts_with('title:') {
					has_title = true
					f := a.split_any(':')
					println('row=${row}; col=${col} title=${f[1]}')
					s.write(row, col, f[1])
					break
				}
			}
			if !has_title {
				s.write(row, col, field.name)
			}
			if row_inc {
				col++
			} else {
				row++
			}
		}
		for d in data {
			println('row=${row}; col=${col}')
			s.write(row, col, d.$(field.name))
			if row_inc {
				col++
			} else {
				row++
			}
		}
		if row_inc {
			row++
			col = col_reset
		} else {
			col++
			row = row_reset
		}
	}
}

// write write `value` to a worksheet cell at (`row`:`col`)
pub fn (s &C.lxw_worksheet) write(row u32, col u16, value All_Data_Type) Lxw_error {
	match value {
		string {
			if value.starts_with('=') {
				return C.worksheet_write_formula(s, row, col, value.str, null_format)
			} else {
				return C.worksheet_write_string(s, row, col, value.str, null_format)
			}
		}
		f64 {
			return C.worksheet_write_number(s, row, col, value, null_format)
		}
		int {
			return C.worksheet_write_number(s, row, col, value, null_format)
		}
		bool {
			return C.worksheet_write_number(s, row, col, int(value), null_format)
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
			return C.worksheet_write_datetime(s, row, col, datetime, null_format)
		}
	}
}

// write_number write a number to a worksheet cell
pub fn (s &C.lxw_worksheet) write_number(row u32, col u16, number f64, format &C.lxw_format) Lxw_error {
	return C.worksheet_write_number(s, row, col, number, format)
}

// write_string write a string to a worksheet cell
pub fn (s &C.lxw_worksheet) write_string(row u32, col u16, str string, format &C.lxw_format) Lxw_error {
	return C.worksheet_write_string(s, row, col, str.str, format)
}

// write_formula write a formula to a worksheet cell
pub fn (s &C.lxw_worksheet) write_formula(row u32, col u16, formula string, format &C.lxw_format) Lxw_error {
	return C.worksheet_write_formula(s, row, col, formula.str, format)
}

// write_array_formula write an array formula to a worksheet cell
pub fn (s &C.lxw_worksheet) write_array_formula(first_row u32, first_col u16, last_row u32, last_col u16, formula string, format &C.lxw_format) Lxw_error {
	return C.worksheet_write_array_formula(s, first_row, first_col, last_row, last_col,
		formula.str, format)
}

// write_dynamic_array_formula write an Excel 365 dynamic array formula to a worksheet range
pub fn (s &C.lxw_worksheet) write_dynamic_array_formula(first_row u32, first_col u16, last_row u32, last_col u16, formula string, format &C.lxw_format) Lxw_error {
	return C.worksheet_write_dynamic_array_formula(s, first_row, first_col, last_row,
		last_col, formula.str, format)
}

// write_dynamic_formula write an Excel 365 dynamic array formula to a worksheet cell
pub fn (s &C.lxw_worksheet) write_dynamic_formula(row u32, col u16, formula string, format &C.lxw_format) Lxw_error {
	return C.worksheet_write_dynamic_formula(s, row, col, formula.str, format)
}

// write_array_formula_num write an array formula with a numerical result to a cell in Excel
pub fn (s &C.lxw_worksheet) write_array_formula_num(first_row u32, first_col u16, last_row u32, last_col u16, formula string, format &C.lxw_format, result f64) Lxw_error {
	return C.worksheet_write_array_formula_num(s, first_row, first_col, last_row, last_col,
		formula.str, format, result)
}

// write_dynamic_array_formula_num write a dynamic array formula with a numerical result to a cell in Excel
pub fn (s &C.lxw_worksheet) write_dynamic_array_formula_num(first_row u32, first_col u16, last_row u32, last_col u16, formula string, format &C.lxw_format, result f64) Lxw_error {
	return C.worksheet_write_dynamic_array_formula_num(s, first_row, first_col, last_row,
		last_col, formula.str, format, result)
}

// write_dynamic_formula_num write a single cell dynamic array formula with a numerical result to a cell
pub fn (s &C.lxw_worksheet) write_dynamic_formula_num(row u32, col u16, formula string, format &C.lxw_format, result f64) Lxw_error {
	return C.worksheet_write_dynamic_formula_num(s, row, col, formula.str, format, result)
}

// write_datetime write a date or time to a worksheet cell
pub fn (s &C.lxw_worksheet) write_datetime(row u32, col u16, datetime &C.lxw_datetime, format &C.lxw_format) Lxw_error {
	return C.worksheet_write_datetime(s, row, col, datetime, format)
}

// write_unixtime write a Unix datetime to a worksheet cell
pub fn (s &C.lxw_worksheet) write_unixtime(row u32, col u16, unixtime i64, format &C.lxw_format) Lxw_error {
	return C.worksheet_write_unixtime(s, row, col, unixtime, format)
}

// write_url write a hyperlink/url to an Excel file
pub fn (s &C.lxw_worksheet) write_url(row u32, col u16, url string, format &C.lxw_format) Lxw_error {
	return C.worksheet_write_url(s, row, col, url.str, format)
}

// write_url_opt write a hyperlink/url to an Excel file
pub fn (s &C.lxw_worksheet) write_url_opt(row_num u32, col_num u16, url string, format &C.lxw_format, string_ &char, tooltip &char) Lxw_error {
	return C.worksheet_write_url_opt(s, row_num, col_num, url.str, format, string_, tooltip)
}

// write_boolean write a formatted boolean worksheet cell
pub fn (s &C.lxw_worksheet) write_boolean(row u32, col u16, value bool, format &C.lxw_format) Lxw_error {
	return C.worksheet_write_boolean(s, row, col, int(value), format)
}

// write_blank write a formatted blank worksheet cell
pub fn (s &C.lxw_worksheet) write_blank(row u32, col u16, format &C.lxw_format) Lxw_error {
	return C.worksheet_write_blank(s, row, col, format)
}

// write_formula_num write a formula to a worksheet cell with a user defined numeric result
pub fn (s &C.lxw_worksheet) write_formula_num(row u32, col u16, formula string, format &C.lxw_format, result f64) Lxw_error {
	return C.worksheet_write_formula_num(s, row, col, formula.str, format, result)
}

// write_formula_str write a formula to a worksheet cell with a user defined string result
pub fn (s &C.lxw_worksheet) write_formula_str(row u32, col u16, formula string, format &C.lxw_format, result &char) Lxw_error {
	return C.worksheet_write_formula_str(s, row, col, formula.str, format, result)
}

// write_rich_string write a "Rich" multi-format string to a worksheet cell
pub fn (s &C.lxw_worksheet) write_rich_string(row u32, col u16, rich_string &&C.lxw_rich_string_tuple, format &C.lxw_format) Lxw_error {
	return C.worksheet_write_rich_string(s, row, col, rich_string, format)
}

// write_comment write a comment to a worksheet cell
pub fn (s &C.lxw_worksheet) write_comment(row u32, col u16, str string) Lxw_error {
	return C.worksheet_write_comment(s, row, col, str.str)
}

// write_comment_opt write a comment to a worksheet cell with options
pub fn (s &C.lxw_worksheet) write_comment_opt(row u32, col u16, str string, options &C.lxw_comment_options) Lxw_error {
	return C.worksheet_write_comment_opt(s, row, col, str.str, options)
}

// set_row set the properties for a row of cells
pub fn (s &C.lxw_worksheet) set_row(row u32, height f64, format &C.lxw_format) Lxw_error {
	return C.worksheet_set_row(s, row, height, format)
}

// set_row_opt set the properties for a row of cells
pub fn (s &C.lxw_worksheet) set_row_opt(row u32, height f64, format &C.lxw_format, options &C.lxw_row_col_options) Lxw_error {
	return C.worksheet_set_row_opt(s, row, height, format, options)
}

// set_row_pixels set the properties for a row of cells, with the height in pixels
pub fn (s &C.lxw_worksheet) set_row_pixels(row u32, pixels u32, format &C.lxw_format) Lxw_error {
	return C.worksheet_set_row_pixels(s, row, pixels, format)
}

// set_row_pixels_opt set the properties for a row of cells, with the height in pixels
pub fn (s &C.lxw_worksheet) set_row_pixels_opt(row u32, pixels u32, format &C.lxw_format, options &C.lxw_row_col_options) Lxw_error {
	return C.worksheet_set_row_pixels_opt(s, row, pixels, format, options)
}

// set_column set the properties for one or more columns of cells
pub fn (s &C.lxw_worksheet) set_column(first_col u16, last_col u16, width f64, format &C.lxw_format) Lxw_error {
	return C.worksheet_set_column(s, first_col, last_col, width, format)
}

// set_column_opt set the properties for one or more columns of cells with options
pub fn (s &C.lxw_worksheet) set_column_opt(first_col u16, last_col u16, width f64, format &C.lxw_format, options &C.lxw_row_col_options) Lxw_error {
	return C.worksheet_set_column_opt(s, first_col, last_col, width, format, options)
}

// set_column_pixels set the properties for one or more columns of cells, with the width in pixels
pub fn (s &C.lxw_worksheet) set_column_pixels(first_col u16, last_col u16, pixels u32, format &C.lxw_format) Lxw_error {
	return C.worksheet_set_column_pixels(s, first_col, last_col, pixels, format)
}

// set_column_pixels_opt set the properties for one or more columns of cells with options, with the width in pixels
pub fn (s &C.lxw_worksheet) set_column_pixels_opt(first_col u16, last_col u16, pixels u32, format &C.lxw_format, options &C.lxw_row_col_options) Lxw_error {
	return C.worksheet_set_column_pixels_opt(s, first_col, last_col, pixels, format, options)
}

// insert_image insert an image in a worksheet cell
pub fn (s &C.lxw_worksheet) insert_image(row u32, col u16, filename string) Lxw_error {
	return C.worksheet_insert_image(s, row, col, filename.str)
}

// insert_image_opt insert an image in a worksheet cell, with options
pub fn (s &C.lxw_worksheet) insert_image_opt(row u32, col u16, filename string, options &C.lxw_image_options) Lxw_error {
	return C.worksheet_insert_image_opt(s, row, col, filename.str, options)
}

// insert_image_buffer insert an image in a worksheet cell, from a memory buffer
pub fn (s &C.lxw_worksheet) insert_image_buffer(row u32, col u16, image_buffer &u8, image_size usize) Lxw_error {
	return C.worksheet_insert_image_buffer(s, row, col, image_buffer, image_size)
}

// insert_image_buffer_opt insert an image in a worksheet cell, from a memory buffer
pub fn (s &C.lxw_worksheet) insert_image_buffer_opt(row u32, col u16, image_buffer &u8, image_size usize, options &C.lxw_image_options) Lxw_error {
	return C.worksheet_insert_image_buffer_opt(s, row, col, image_buffer, image_size,
		options)
}

// set_background set the background image for a worksheet
pub fn (s &C.lxw_worksheet) set_background(filename string) Lxw_error {
	return C.worksheet_set_background(s, filename.str)
}

// set_background_buffer set the background image for a worksheet, from a buffer
pub fn (s &C.lxw_worksheet) set_background_buffer(image_buffer &u8, image_size usize) Lxw_error {
	return C.worksheet_set_background_buffer(s, image_buffer, image_size)
}

// insert_chart insert a chart object into a worksheet
pub fn (s &C.lxw_worksheet) insert_chart(row u32, col u16, chart &C.lxw_chart) Lxw_error {
	return C.worksheet_insert_chart(s, row, col, chart)
}

// insert_chart_opt insert a chart object into a worksheet, with options
pub fn (s &C.lxw_worksheet) insert_chart_opt(row u32, col u16, chart &C.lxw_chart, user_options &C.lxw_chart_options) Lxw_error {
	return C.worksheet_insert_chart_opt(s, row, col, chart, user_options)
}

// merge_range merge a range of cells
pub fn (s &C.lxw_worksheet) merge_range(first_row u32, first_col u16, last_row u32, last_col u16, str string, format &C.lxw_format) Lxw_error {
	return C.worksheet_merge_range(s, first_row, first_col, last_row, last_col, str.str,
		format)
}

// autofilter set the autofilter area in the worksheet
pub fn (s &C.lxw_worksheet) autofilter(first_row u32, first_col u16, last_row u32, last_col u16) Lxw_error {
	return C.worksheet_autofilter(s, first_row, first_col, last_row, last_col)
}

// filter_column write a filter rule to an autofilter column
pub fn (s &C.lxw_worksheet) filter_column(col u16, rule &C.lxw_filter_rule) Lxw_error {
	return C.worksheet_filter_column(s, col, rule)
}

// filter_column2 write two filter rules to an autofilter column
pub fn (s &C.lxw_worksheet) filter_column2(col u16, rule1 &C.lxw_filter_rule, rule2 &C.lxw_filter_rule, and_or u8) Lxw_error {
	return C.worksheet_filter_column2(s, col, rule1, rule2, and_or)
}

// filter_list write multiple string filters to an autofilter column
pub fn (s &C.lxw_worksheet) filter_list(col u16, list &&u8) Lxw_error {
	return C.worksheet_filter_list(s, col, list)
}

// data_validation_cell add a data validation to a cell
pub fn (s &C.lxw_worksheet) data_validation_cell(row u32, col u16, validation &C.lxw_data_validation) Lxw_error {
	return C.worksheet_data_validation_cell(s, row, col, validation)
}

// data_validation_range add a data validation to a range
pub fn (s &C.lxw_worksheet) data_validation_range(first_row u32, first_col u16, last_row u32, last_col u16, validation &C.lxw_data_validation) Lxw_error {
	return C.worksheet_data_validation_range(s, first_row, first_col, last_row, last_col,
		validation)
}

// conditional_format_cell add a conditional format to a worksheet cell
pub fn (s &C.lxw_worksheet) conditional_format_cell(row u32, col u16, conditional_format &C.lxw_conditional_format) Lxw_error {
	return C.worksheet_conditional_format_cell(s, row, col, conditional_format)
}

// conditional_format_range add a conditional format to a worksheet range
pub fn (s &C.lxw_worksheet) conditional_format_range(first_row u32, first_col u16, last_row u32, last_col u16, conditional_format &C.lxw_conditional_format) Lxw_error {
	return C.worksheet_conditional_format_range(s, first_row, first_col, last_row, last_col,
		conditional_format)
}

// insert_button insert a button object into a worksheet
pub fn (s &C.lxw_worksheet) insert_button(row u32, col u16, options &C.lxw_button_options) Lxw_error {
	return C.worksheet_insert_button(s, row, col, options)
}

// add_table add an Excel table to a worksheet
pub fn (s &C.lxw_worksheet) add_table(first_row u32, first_col u16, last_row u32, last_col u16, options &C.lxw_table_options) Lxw_error {
	return C.worksheet_add_table(s, first_row, first_col, last_row, last_col, options)
}

// activate make a worksheet the active, i.e., visible worksheet
pub fn (s &C.lxw_worksheet) activate() &C.lxw_worksheet {
	C.worksheet_activate(s)
	return unsafe { s }
}

// @select set a worksheet tab as selected
pub fn (s &C.lxw_worksheet) @select() &C.lxw_worksheet {
	C.worksheet_select(s)
	return unsafe { s }
}

// hide hide the current worksheet
pub fn (s &C.lxw_worksheet) hide() &C.lxw_worksheet {
	C.worksheet_hide(s)
	return unsafe { s }
}

// set_first_sheet set current worksheet as the first visible sheet tab
pub fn (s &C.lxw_worksheet) set_first_sheet() &C.lxw_worksheet {
	C.worksheet_set_first_sheet(s)
	return unsafe { s }
}

// freeze_panes split and freeze a worksheet into panes
pub fn (s &C.lxw_worksheet) freeze_panes(row u32, col u16) &C.lxw_worksheet {
	C.worksheet_freeze_panes(s, row, col)
	return unsafe { s }
}

// split_panes split a worksheet into panes
pub fn (s &C.lxw_worksheet) split_panes(vertical f64, horizontal f64) &C.lxw_worksheet {
	C.worksheet_split_panes(s, vertical, horizontal)
	return unsafe { s }
}

// freeze_panes_opt set panes and mark them as frozen. With extra options
pub fn (s &C.lxw_worksheet) freeze_panes_opt(first_row u32, first_col u16, top_row u32, left_col u16, type_ u8) &C.lxw_worksheet {
	C.worksheet_freeze_panes_opt(s, first_row, first_col, top_row, left_col, type_)
	return unsafe { s }
}

// split_panes_opt set panes and mark them as split.With extra options
pub fn (s &C.lxw_worksheet) split_panes_opt(vertical f64, horizontal f64, top_row u32, left_col u16) &C.lxw_worksheet {
	C.worksheet_split_panes_opt(s, vertical, horizontal, top_row, left_col)
	return unsafe { s }
}

// set_selection set the selected cell or cells in a worksheet
pub fn (s &C.lxw_worksheet) set_selection(first_row u32, first_col u16, last_row u32, last_col u16) &C.lxw_worksheet {
	C.worksheet_set_selection(s, first_row, first_col, last_row, last_col)
	return unsafe { s }
}

// set_top_left_cell set the first visible cell at the top left of a worksheet
pub fn (s &C.lxw_worksheet) set_top_left_cell(row u32, col u16) &C.lxw_worksheet {
	C.worksheet_set_top_left_cell(s, row, col)
	return unsafe { s }
}

// set_landscape set the page orientation as landscape
pub fn (s &C.lxw_worksheet) set_landscape() &C.lxw_worksheet {
	C.worksheet_set_landscape(s)
	return unsafe { s }
}

// set_portrait set the page orientation as portrait
pub fn (s &C.lxw_worksheet) set_portrait() &C.lxw_worksheet {
	C.worksheet_set_portrait(s)
	return unsafe { s }
}

// set_page_view set the page layout to page view mode
pub fn (s &C.lxw_worksheet) set_page_view() &C.lxw_worksheet {
	C.worksheet_set_page_view(s)
	return unsafe { s }
}

// set_paper set the paper type for printing
pub fn (s &C.lxw_worksheet) set_paper(paper_type Paper_format_type) &C.lxw_worksheet {
	C.worksheet_set_paper(s, u8(paper_type))
	return unsafe { s }
}

// set_margins set the worksheet margins for the printed page
pub fn (s &C.lxw_worksheet) set_margins(left f64, right f64, top f64, bottom f64) &C.lxw_worksheet {
	C.worksheet_set_margins(s, left, right, top, bottom)
	return unsafe { s }
}

// set_header set the printed page header caption
pub fn (s &C.lxw_worksheet) set_header(str string) Lxw_error {
	return C.worksheet_set_header(s, str.str)
}

// set_footer set the printed page footer caption
pub fn (s &C.lxw_worksheet) set_footer(str string) Lxw_error {
	return C.worksheet_set_footer(s, str.str)
}

// set_header_opt set the printed page header caption with additional options
pub fn (s &C.lxw_worksheet) set_header_opt(str string, options &C.lxw_header_footer_options) Lxw_error {
	return C.worksheet_set_header_opt(s, str.str, options)
}

// set_footer_opt set the printed page footer caption with additional options
pub fn (s &C.lxw_worksheet) set_footer_opt(str string, options &C.lxw_header_footer_options) Lxw_error {
	return C.worksheet_set_footer_opt(s, str.str, options)
}

// set_h_pagebreaks set the horizontal page breaks on a worksheet
pub fn (s &C.lxw_worksheet) set_h_pagebreaks(breaks &u32) Lxw_error {
	return C.worksheet_set_h_pagebreaks(s, breaks)
}

// set_v_pagebreaks set the vertical page breaks on a worksheet
pub fn (s &C.lxw_worksheet) set_v_pagebreaks(breaks &u16) Lxw_error {
	return C.worksheet_set_v_pagebreaks(s, breaks)
}

// print_across set the order in which pages are printed
pub fn (s &C.lxw_worksheet) print_across() &C.lxw_worksheet {
	C.worksheet_print_across(s)
	return unsafe { s }
}

// set_zoom set the worksheet zoom factor
pub fn (s &C.lxw_worksheet) set_zoom(scale u16) &C.lxw_worksheet {
	C.worksheet_set_zoom(s, scale)
	return unsafe { s }
}

// gridlines set the option to display or hide gridlines on the screen and the printed page
pub fn (s &C.lxw_worksheet) gridlines(option Lxw_gridlines) &C.lxw_worksheet {
	C.worksheet_gridlines(s, u8(option))
	return unsafe { s }
}

// center_horizontally center the printed page horizontally
pub fn (s &C.lxw_worksheet) center_horizontally() &C.lxw_worksheet {
	C.worksheet_center_horizontally(s)
	return unsafe { s }
}

// center_vertically center the printed page vertically
pub fn (s &C.lxw_worksheet) center_vertically() &C.lxw_worksheet {
	C.worksheet_center_vertically(s)
	return unsafe { s }
}

// print_row_col_headers set the option to print the row and column headers on the printed page
pub fn (s &C.lxw_worksheet) print_row_col_headers() &C.lxw_worksheet {
	C.worksheet_print_row_col_headers(s)
	return unsafe { s }
}

// repeat_rows set the number of rows to repeat at the top of each printed page
pub fn (s &C.lxw_worksheet) repeat_rows(first_row u32, last_row u32) Lxw_error {
	return C.worksheet_repeat_rows(s, first_row, last_row)
}

// repeat_columns set the number of columns to repeat at the top of each printed page
pub fn (s &C.lxw_worksheet) repeat_columns(first_col u16, last_col u16) Lxw_error {
	return C.worksheet_repeat_columns(s, first_col, last_col)
}

// print_area set the print area for a worksheet
pub fn (s &C.lxw_worksheet) print_area(first_row u32, first_col u16, last_row u32, last_col u16) Lxw_error {
	return C.worksheet_print_area(s, first_row, first_col, last_row, last_col)
}

// fit_to_pages fit the printed area to a specific number of pages both vertically and horizontally
pub fn (s &C.lxw_worksheet) fit_to_pages(width u16, height u16) &C.lxw_worksheet {
	C.worksheet_fit_to_pages(s, width, height)
	return unsafe { s }
}

// set_start_page set the start/first page number when printing
pub fn (s &C.lxw_worksheet) set_start_page(start_page u16) &C.lxw_worksheet {
	C.worksheet_set_start_page(s, start_page)
	return unsafe { s }
}

// set_print_scale set the scale factor for the printed page
pub fn (s &C.lxw_worksheet) set_print_scale(scale u16) &C.lxw_worksheet {
	C.worksheet_set_print_scale(s, scale)
	return unsafe { s }
}

// print_black_and_white set the worksheet to print in black and white
pub fn (s &C.lxw_worksheet) print_black_and_white() &C.lxw_worksheet {
	C.worksheet_print_black_and_white(s)
	return unsafe { s }
}

// right_to_left display the worksheet cells from right to left for some versions of Excel
pub fn (s &C.lxw_worksheet) right_to_left() &C.lxw_worksheet {
	C.worksheet_right_to_left(s)
	return unsafe { s }
}

// hide_zero hide zero values in worksheet cells
pub fn (s &C.lxw_worksheet) hide_zero() &C.lxw_worksheet {
	C.worksheet_hide_zero(s)
	return unsafe { s }
}

// set_tab_color set the color of the worksheet tab
pub fn (s &C.lxw_worksheet) set_tab_color(color u32) &C.lxw_worksheet {
	C.worksheet_set_tab_color(s, color)
	return unsafe { s }
}

// protect protect elements of a worksheet from modification
pub fn (s &C.lxw_worksheet) protect(password string, options &C.lxw_protection) &C.lxw_worksheet {
	C.worksheet_protect(s, password.str, options)
	return unsafe { s }
}

// outline_settings set the Outline and Grouping display properties
pub fn (s &C.lxw_worksheet) outline_settings(visible u8, symbols_below u8, symbols_right u8, auto_style u8) &C.lxw_worksheet {
	C.worksheet_outline_settings(s, visible, symbols_below, symbols_right, auto_style)
	return unsafe { s }
}

// set_default_row set the default row properties
pub fn (s &C.lxw_worksheet) set_default_row(height f64, hide_unused_rows u8) &C.lxw_worksheet {
	C.worksheet_set_default_row(s, height, hide_unused_rows)
	return unsafe { s }
}

// set_vba_name set the VBA name for the worksheet
pub fn (s &C.lxw_worksheet) set_vba_name(name string) Lxw_error {
	return C.worksheet_set_vba_name(s, name.str)
}

// show_comments make all comments in the worksheet visible
pub fn (s &C.lxw_worksheet) show_comments() &C.lxw_worksheet {
	C.worksheet_show_comments(s)
	return unsafe { s }
}

// set_comments_author set the default author of the cell comments
pub fn (s &C.lxw_worksheet) set_comments_author(author string) &C.lxw_worksheet {
	C.worksheet_set_comments_author(s, author.str)
	return unsafe { s }
}

// ignore_errors ignore various Excel errors/warnings in a worksheet for user defined ranges
pub fn (s &C.lxw_worksheet) ignore_errors(type_ u8, range string) Lxw_error {
	return C.worksheet_ignore_errors(s, type_, range.str)
}
