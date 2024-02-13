# module vxlsx


## Contents
- [Constants](#Constants)
- [lxw_datetime_to_excel_datetime](#lxw_datetime_to_excel_datetime)
- [lxw_strerror](#lxw_strerror)
- [lxw_unixtime_to_excel_date](#lxw_unixtime_to_excel_date)
- [lxw_version](#lxw_version)
- [lxw_version_id](#lxw_version_id)
- [new_workbook](#new_workbook)
- [new_workbook_opt](#new_workbook_opt)
- [All_Data_Type](#All_Data_Type)
- [C.lxw_chart](#C.lxw_chart)
  - [add_series](#add_series)
  - [get](#get)
  - [title_set_name](#title_set_name)
  - [title_set_name_range](#title_set_name_range)
  - [title_set_name_font](#title_set_name_font)
  - [title_off](#title_off)
  - [legend_set_position](#legend_set_position)
  - [legend_set_font](#legend_set_font)
  - [legend_delete_series](#legend_delete_series)
  - [chartarea_set_line](#chartarea_set_line)
  - [chartarea_set_fill](#chartarea_set_fill)
  - [chartarea_set_pattern](#chartarea_set_pattern)
  - [plotarea_set_line](#plotarea_set_line)
  - [plotarea_set_fill](#plotarea_set_fill)
  - [plotarea_set_pattern](#plotarea_set_pattern)
  - [set_style](#set_style)
  - [set_table](#set_table)
  - [set_table_grid](#set_table_grid)
  - [set_table_font](#set_table_font)
  - [set_up_down_bars](#set_up_down_bars)
  - [set_up_down_bars_format](#set_up_down_bars_format)
  - [set_drop_lines](#set_drop_lines)
  - [set_high_low_lines](#set_high_low_lines)
  - [set_series_overlap](#set_series_overlap)
  - [set_series_gap](#set_series_gap)
  - [show_blanks_as](#show_blanks_as)
  - [show_hidden_data](#show_hidden_data)
  - [set_rotation](#set_rotation)
  - [set_hole_size](#set_hole_size)
- [C.lxw_chart_axis](#C.lxw_chart_axis)
  - [set_name](#set_name)
  - [set_name_range](#set_name_range)
  - [set_name_font](#set_name_font)
  - [set_num_font](#set_num_font)
  - [set_num_format](#set_num_format)
  - [set_line](#set_line)
  - [set_fill](#set_fill)
  - [set_pattern](#set_pattern)
  - [set_reverse](#set_reverse)
  - [set_crossing](#set_crossing)
  - [set_crossing_max](#set_crossing_max)
  - [set_crossing_min](#set_crossing_min)
  - [off](#off)
  - [set_position](#set_position)
  - [set_label_position](#set_label_position)
  - [set_label_align](#set_label_align)
  - [set_min](#set_min)
  - [set_max](#set_max)
  - [set_log_base](#set_log_base)
  - [set_major_tick_mark](#set_major_tick_mark)
  - [set_minor_tick_mark](#set_minor_tick_mark)
  - [set_interval_unit](#set_interval_unit)
  - [set_interval_tick](#set_interval_tick)
  - [set_major_unit](#set_major_unit)
  - [set_minor_unit](#set_minor_unit)
  - [set_display_units](#set_display_units)
  - [set_display_units_visible](#set_display_units_visible)
  - [major_gridlines_set_visible](#major_gridlines_set_visible)
  - [minor_gridlines_set_visible](#minor_gridlines_set_visible)
  - [major_gridlines_set_line](#major_gridlines_set_line)
  - [minor_gridlines_set_line](#minor_gridlines_set_line)
- [C.lxw_chart_series](#C.lxw_chart_series)
  - [set_categories](#set_categories)
  - [set_values](#set_values)
  - [set_name](#set_name)
  - [set_name_range](#set_name_range)
  - [set_line](#set_line)
  - [set_fill](#set_fill)
  - [set_invert_if_negative](#set_invert_if_negative)
  - [set_pattern](#set_pattern)
  - [set_marker_type](#set_marker_type)
  - [set_marker_size](#set_marker_size)
  - [set_marker_line](#set_marker_line)
  - [set_marker_fill](#set_marker_fill)
  - [set_marker_pattern](#set_marker_pattern)
  - [set_points](#set_points)
  - [set_smooth](#set_smooth)
  - [set_labels](#set_labels)
  - [set_labels_options](#set_labels_options)
  - [set_labels_custom](#set_labels_custom)
  - [set_labels_separator](#set_labels_separator)
  - [set_labels_position](#set_labels_position)
  - [set_labels_leader_line](#set_labels_leader_line)
  - [set_labels_legend](#set_labels_legend)
  - [set_labels_percentage](#set_labels_percentage)
  - [set_labels_num_format](#set_labels_num_format)
  - [set_labels_font](#set_labels_font)
  - [set_labels_line](#set_labels_line)
  - [set_labels_fill](#set_labels_fill)
  - [set_labels_pattern](#set_labels_pattern)
  - [set_trendline](#set_trendline)
  - [set_trendline_forecast](#set_trendline_forecast)
  - [set_trendline_equation](#set_trendline_equation)
  - [set_trendline_r_squared](#set_trendline_r_squared)
  - [set_trendline_intercept](#set_trendline_intercept)
  - [set_trendline_name](#set_trendline_name)
  - [set_trendline_line](#set_trendline_line)
  - [get_error_bars](#get_error_bars)
- [C.lxw_chartsheet](#C.lxw_chartsheet)
  - [set_chart](#set_chart)
  - [activate](#activate)
  - [@select](#@select)
  - [hide](#hide)
  - [set_first_sheet](#set_first_sheet)
  - [set_tab_color](#set_tab_color)
  - [protect](#protect)
  - [set_zoom](#set_zoom)
  - [set_landscape](#set_landscape)
  - [set_portrait](#set_portrait)
  - [set_paper](#set_paper)
  - [set_margins](#set_margins)
  - [set_header](#set_header)
  - [set_footer](#set_footer)
  - [set_header_opt](#set_header_opt)
  - [set_footer_opt](#set_footer_opt)
- [C.lxw_format](#C.lxw_format)
  - [set_font_name](#set_font_name)
  - [set_font_size](#set_font_size)
  - [set_font_color](#set_font_color)
  - [set_bold](#set_bold)
  - [set_italic](#set_italic)
  - [set_underline](#set_underline)
  - [set_font_strikeout](#set_font_strikeout)
  - [set_font_script](#set_font_script)
  - [set_num_format](#set_num_format)
  - [set_num_format_index](#set_num_format_index)
  - [set_unlocked](#set_unlocked)
  - [set_hidden](#set_hidden)
  - [set_align](#set_align)
  - [set_text_wrap](#set_text_wrap)
  - [set_rotation](#set_rotation)
  - [set_indent](#set_indent)
  - [set_shrink](#set_shrink)
  - [set_pattern](#set_pattern)
  - [set_bg_color](#set_bg_color)
  - [set_fg_color](#set_fg_color)
  - [set_border](#set_border)
  - [set_bottom](#set_bottom)
  - [set_top](#set_top)
  - [set_left](#set_left)
  - [set_right](#set_right)
  - [set_border_color](#set_border_color)
  - [set_bottom_color](#set_bottom_color)
  - [set_top_color](#set_top_color)
  - [set_left_color](#set_left_color)
  - [set_right_color](#set_right_color)
  - [set_diag_type](#set_diag_type)
  - [set_diag_border](#set_diag_border)
  - [set_diag_color](#set_diag_color)
  - [set_quote_prefix](#set_quote_prefix)
  - [set_font_outline](#set_font_outline)
  - [set_font_shadow](#set_font_shadow)
  - [set_font_family](#set_font_family)
  - [set_font_charset](#set_font_charset)
  - [set_font_scheme](#set_font_scheme)
  - [set_font_condense](#set_font_condense)
  - [set_font_extend](#set_font_extend)
  - [set_reading_order](#set_reading_order)
  - [set_theme](#set_theme)
  - [set_hyperlink](#set_hyperlink)
  - [set_color_indexed](#set_color_indexed)
  - [set_font_only](#set_font_only)
- [C.lxw_series_error_bars](#C.lxw_series_error_bars)
  - [set](#set)
  - [direction](#direction)
  - [endcap](#endcap)
  - [line](#line)
- [C.lxw_workbook](#C.lxw_workbook)
  - [add_worksheet](#add_worksheet)
  - [add_chartsheet](#add_chartsheet)
  - [add_format](#add_format)
  - [add_chart](#add_chart)
  - [close](#close)
  - [set_properties](#set_properties)
  - [set_custom_property](#set_custom_property)
  - [define_name](#define_name)
  - [get_default_url_format](#get_default_url_format)
  - [get_worksheet_by_name](#get_worksheet_by_name)
  - [get_chartsheet_by_name](#get_chartsheet_by_name)
  - [validate_sheet_name](#validate_sheet_name)
  - [add_vba_project](#add_vba_project)
  - [add_signed_vba_project](#add_signed_vba_project)
  - [set_vba_name](#set_vba_name)
  - [read_only_recommended](#read_only_recommended)
  - [unset_default_url_format](#unset_default_url_format)
- [C.lxw_worksheet](#C.lxw_worksheet)
  - [write_struct](#write_struct)
  - [write](#write)
  - [write_number](#write_number)
  - [write_string](#write_string)
  - [write_formula](#write_formula)
  - [write_array_formula](#write_array_formula)
  - [write_dynamic_array_formula](#write_dynamic_array_formula)
  - [write_dynamic_formula](#write_dynamic_formula)
  - [write_array_formula_num](#write_array_formula_num)
  - [write_dynamic_array_formula_num](#write_dynamic_array_formula_num)
  - [write_dynamic_formula_num](#write_dynamic_formula_num)
  - [write_datetime](#write_datetime)
  - [write_unixtime](#write_unixtime)
  - [write_url](#write_url)
  - [write_url_opt](#write_url_opt)
  - [write_boolean](#write_boolean)
  - [write_blank](#write_blank)
  - [write_formula_num](#write_formula_num)
  - [write_formula_str](#write_formula_str)
  - [write_rich_string](#write_rich_string)
  - [write_comment](#write_comment)
  - [write_comment_opt](#write_comment_opt)
  - [set_row](#set_row)
  - [set_row_opt](#set_row_opt)
  - [set_row_pixels](#set_row_pixels)
  - [set_row_pixels_opt](#set_row_pixels_opt)
  - [set_column](#set_column)
  - [set_column_opt](#set_column_opt)
  - [set_column_pixels](#set_column_pixels)
  - [set_column_pixels_opt](#set_column_pixels_opt)
  - [insert_image](#insert_image)
  - [insert_image_opt](#insert_image_opt)
  - [insert_image_buffer](#insert_image_buffer)
  - [insert_image_buffer_opt](#insert_image_buffer_opt)
  - [set_background](#set_background)
  - [set_background_buffer](#set_background_buffer)
  - [insert_chart](#insert_chart)
  - [insert_chart_opt](#insert_chart_opt)
  - [merge_range](#merge_range)
  - [autofilter](#autofilter)
  - [filter_column](#filter_column)
  - [filter_column2](#filter_column2)
  - [filter_list](#filter_list)
  - [data_validation_cell](#data_validation_cell)
  - [data_validation_range](#data_validation_range)
  - [conditional_format_cell](#conditional_format_cell)
  - [conditional_format_range](#conditional_format_range)
  - [insert_button](#insert_button)
  - [add_table](#add_table)
  - [activate](#activate)
  - [@select](#@select)
  - [hide](#hide)
  - [set_first_sheet](#set_first_sheet)
  - [freeze_panes](#freeze_panes)
  - [split_panes](#split_panes)
  - [freeze_panes_opt](#freeze_panes_opt)
  - [split_panes_opt](#split_panes_opt)
  - [set_selection](#set_selection)
  - [set_top_left_cell](#set_top_left_cell)
  - [set_landscape](#set_landscape)
  - [set_portrait](#set_portrait)
  - [set_page_view](#set_page_view)
  - [set_paper](#set_paper)
  - [set_margins](#set_margins)
  - [set_header](#set_header)
  - [set_footer](#set_footer)
  - [set_header_opt](#set_header_opt)
  - [set_footer_opt](#set_footer_opt)
  - [set_h_pagebreaks](#set_h_pagebreaks)
  - [set_v_pagebreaks](#set_v_pagebreaks)
  - [print_across](#print_across)
  - [set_zoom](#set_zoom)
  - [gridlines](#gridlines)
  - [center_horizontally](#center_horizontally)
  - [center_vertically](#center_vertically)
  - [print_row_col_headers](#print_row_col_headers)
  - [repeat_rows](#repeat_rows)
  - [repeat_columns](#repeat_columns)
  - [print_area](#print_area)
  - [fit_to_pages](#fit_to_pages)
  - [set_start_page](#set_start_page)
  - [set_print_scale](#set_print_scale)
  - [print_black_and_white](#print_black_and_white)
  - [right_to_left](#right_to_left)
  - [hide_zero](#hide_zero)
  - [set_tab_color](#set_tab_color)
  - [protect](#protect)
  - [outline_settings](#outline_settings)
  - [set_default_row](#set_default_row)
  - [set_vba_name](#set_vba_name)
  - [show_comments](#show_comments)
  - [set_comments_author](#set_comments_author)
  - [ignore_errors](#ignore_errors)
- [Cell_types](#Cell_types)
- [Lxw_chart_axis_display_unit](#Lxw_chart_axis_display_unit)
- [Lxw_chart_axis_label_alignment](#Lxw_chart_axis_label_alignment)
- [Lxw_chart_axis_label_position](#Lxw_chart_axis_label_position)
- [Lxw_chart_axis_tick_position](#Lxw_chart_axis_tick_position)
- [Lxw_chart_axis_type](#Lxw_chart_axis_type)
- [Lxw_chart_blank](#Lxw_chart_blank)
- [Lxw_chart_error_bar_axis](#Lxw_chart_error_bar_axis)
- [Lxw_chart_error_bar_cap](#Lxw_chart_error_bar_cap)
- [Lxw_chart_error_bar_direction](#Lxw_chart_error_bar_direction)
- [Lxw_chart_error_bar_type](#Lxw_chart_error_bar_type)
- [Lxw_chart_grouping](#Lxw_chart_grouping)
- [Lxw_chart_label_position](#Lxw_chart_label_position)
- [Lxw_chart_label_separator](#Lxw_chart_label_separator)
- [Lxw_chart_legend_position](#Lxw_chart_legend_position)
- [Lxw_chart_line_dash_type](#Lxw_chart_line_dash_type)
- [Lxw_chart_marker_type](#Lxw_chart_marker_type)
- [Lxw_chart_pattern_type](#Lxw_chart_pattern_type)
- [Lxw_chart_subtype](#Lxw_chart_subtype)
- [Lxw_chart_tick_mark](#Lxw_chart_tick_mark)
- [Lxw_chart_trendline_type](#Lxw_chart_trendline_type)
- [Lxw_chart_type](#Lxw_chart_type)
- [Lxw_comment_display_types](#Lxw_comment_display_types)
- [Lxw_conditional_bar_axis_position](#Lxw_conditional_bar_axis_position)
- [Lxw_conditional_criteria](#Lxw_conditional_criteria)
- [Lxw_conditional_format_bar_direction](#Lxw_conditional_format_bar_direction)
- [Lxw_conditional_format_rule_types](#Lxw_conditional_format_rule_types)
- [Lxw_conditional_format_types](#Lxw_conditional_format_types)
- [Lxw_conditional_icon_types](#Lxw_conditional_icon_types)
- [Lxw_defined_colors](#Lxw_defined_colors)
- [Lxw_error](#Lxw_error)
- [Lxw_filter_criteria](#Lxw_filter_criteria)
- [Lxw_filter_operator](#Lxw_filter_operator)
- [Lxw_filter_type](#Lxw_filter_type)
- [Lxw_format_alignments](#Lxw_format_alignments)
- [Lxw_format_borders](#Lxw_format_borders)
- [Lxw_format_diagonal_types](#Lxw_format_diagonal_types)
- [Lxw_format_patterns](#Lxw_format_patterns)
- [Lxw_format_scripts](#Lxw_format_scripts)
- [Lxw_format_underlines](#Lxw_format_underlines)
- [Lxw_gridlines](#Lxw_gridlines)
- [Lxw_ignore_errors](#Lxw_ignore_errors)
- [Lxw_image_position](#Lxw_image_position)
- [Lxw_object_position](#Lxw_object_position)
- [Lxw_table_style_type](#Lxw_table_style_type)
- [Lxw_table_total_functions](#Lxw_table_total_functions)
- [Lxw_validation_boolean](#Lxw_validation_boolean)
- [Lxw_validation_criteria](#Lxw_validation_criteria)
- [Lxw_validation_error_types](#Lxw_validation_error_types)
- [Lxw_validation_types](#Lxw_validation_types)
- [Pane_types](#Pane_types)
- [Paper_format_type](#Paper_format_type)
- [C.lxw_button_options](#C.lxw_button_options)
- [C.lxw_chart_fill](#C.lxw_chart_fill)
- [C.lxw_chart_line](#C.lxw_chart_line)
- [C.lxw_chart_options](#C.lxw_chart_options)
- [C.lxw_chart_pattern](#C.lxw_chart_pattern)
- [C.lxw_comment_options](#C.lxw_comment_options)
- [C.lxw_doc_properties](#C.lxw_doc_properties)
- [C.lxw_header_footer_options](#C.lxw_header_footer_options)
- [C.lxw_image_options](#C.lxw_image_options)
- [C.lxw_protection](#C.lxw_protection)
- [C.lxw_row_col_options](#C.lxw_row_col_options)
- [C.lxw_table_options](#C.lxw_table_options)
- [C.lxw_workbook_options](#C.lxw_workbook_options)
- [WorkBook_Options](#WorkBook_Options)

## Constants
[[Return to contents]](#Contents)

## lxw_datetime_to_excel_datetime
```v
fn lxw_datetime_to_excel_datetime(value time.Time) f64
```
lxw_datetime_to_excel_datetime converts a lxw_datetime to an Excel datetime number

[[Return to contents]](#Contents)

## lxw_strerror
```v
fn lxw_strerror(error_num Lxw_error) string
```
lxw_strerror converts a libxlsxwriter error number to a string

[[Return to contents]](#Contents)

## lxw_unixtime_to_excel_date
```v
fn lxw_unixtime_to_excel_date(unixtime i64) f64
```
lxw_unixtime_to_excel_date converts a unix datetime to an Excel datetime number

[[Return to contents]](#Contents)

## lxw_version
```v
fn lxw_version() string
```
lxw_version retrieve the library version

[[Return to contents]](#Contents)

## lxw_version_id
```v
fn lxw_version_id() u16
```
lxw_version_id retrieve the library version ID

[[Return to contents]](#Contents)

## new_workbook
```v
fn new_workbook(filename string) &C.lxw_workbook
```
new_workbook create a new workbook object

[[Return to contents]](#Contents)

## new_workbook_opt
```v
fn new_workbook_opt(filename string, options &WorkBook_Options) &C.lxw_workbook
```
new_workbook_opt create a new workbook object, and set the workbook options

[[Return to contents]](#Contents)

## All_Data_Type
[[Return to contents]](#Contents)

## C.lxw_chart
## add_series
```v
fn (c &C.lxw_chart) add_series(categories string, values string) &C.lxw_chart_series
```
add_series add a data series to a chart

[[Return to contents]](#Contents)

## get
```v
fn (c &C.lxw_chart) get(axis_type Lxw_chart_axis_type) &C.lxw_chart_axis
```
get get an axis pointer from a chart

[[Return to contents]](#Contents)

## title_set_name
```v
fn (c &C.lxw_chart) title_set_name(name string) &C.lxw_chart
```
title_set_name set the title of the chart

[[Return to contents]](#Contents)

## title_set_name_range
```v
fn (c &C.lxw_chart) title_set_name_range(sheetname string, row u32, col u16) &C.lxw_chart
```
title_set_name_range set a chart title formula using row and column values

[[Return to contents]](#Contents)

## title_set_name_font
```v
fn (c &C.lxw_chart) title_set_name_font(font &C.lxw_chart_font) &C.lxw_chart
```
title_set_name_font set the font properties for a chart title

[[Return to contents]](#Contents)

## title_off
```v
fn (c &C.lxw_chart) title_off() &C.lxw_chart
```
title_off turn off an automatic chart title

[[Return to contents]](#Contents)

## legend_set_position
```v
fn (c &C.lxw_chart) legend_set_position(position Lxw_chart_legend_position) &C.lxw_chart
```
legend_set_position set the position of the chart legend

[[Return to contents]](#Contents)

## legend_set_font
```v
fn (c &C.lxw_chart) legend_set_font(font &C.lxw_chart_font) &C.lxw_chart
```
legend_set_font set the font properties for a chart legend

[[Return to contents]](#Contents)

## legend_delete_series
```v
fn (c &C.lxw_chart) legend_delete_series(delete_series &i16) Lxw_error
```
legend_delete_series remove one or more series from the the legend

[[Return to contents]](#Contents)

## chartarea_set_line
```v
fn (c &C.lxw_chart) chartarea_set_line(line &C.lxw_chart_line) &C.lxw_chart
```
chartarea_set_line set the line properties for a chartarea

[[Return to contents]](#Contents)

## chartarea_set_fill
```v
fn (c &C.lxw_chart) chartarea_set_fill(fill &C.lxw_chart_fill) &C.lxw_chart
```
chartarea_set_fil set the fill properties for a chartareal

[[Return to contents]](#Contents)

## chartarea_set_pattern
```v
fn (c &C.lxw_chart) chartarea_set_pattern(pattern &C.lxw_chart_pattern) &C.lxw_chart
```
chartarea_set_pattern set the pattern properties for a chartarea

[[Return to contents]](#Contents)

## plotarea_set_line
```v
fn (c &C.lxw_chart) plotarea_set_line(line &C.lxw_chart_line) &C.lxw_chart
```
plotarea_set_line set the line properties for a plotarea

[[Return to contents]](#Contents)

## plotarea_set_fill
```v
fn (c &C.lxw_chart) plotarea_set_fill(fill &C.lxw_chart_fill) &C.lxw_chart
```
plotarea_set_fill set the fill properties for a plotarea

[[Return to contents]](#Contents)

## plotarea_set_pattern
```v
fn (c &C.lxw_chart) plotarea_set_pattern(pattern &C.lxw_chart_pattern) &C.lxw_chart
```
plotarea_set_pattern set the pattern properties for a plotarea

[[Return to contents]](#Contents)

## set_style
```v
fn (c &C.lxw_chart) set_style(style_id u8) &C.lxw_chart
```
set_style set the chart style type, `style_id` is an index representing the chart style, 1 - 48

[[Return to contents]](#Contents)

## set_table
```v
fn (c &C.lxw_chart) set_table() &C.lxw_chart
```
set_table turn on a data table below the horizontal axis

[[Return to contents]](#Contents)

## set_table_grid
```v
fn (c &C.lxw_chart) set_table_grid(horizontal bool, vertical bool, outline bool, legend_keys bool) &C.lxw_chart
```
set_table_grid turn on/off grid options for a chart data table

[[Return to contents]](#Contents)

## set_table_font
```v
fn (c &C.lxw_chart) set_table_font(font &C.lxw_chart_font) &C.lxw_chart
```
set_table_font

[[Return to contents]](#Contents)

## set_up_down_bars
```v
fn (c &C.lxw_chart) set_up_down_bars() &C.lxw_chart
```
set_up_down_bars turn on up-down bars for the chart

[[Return to contents]](#Contents)

## set_up_down_bars_format
```v
fn (c &C.lxw_chart) set_up_down_bars_format(up_bar_line &C.lxw_chart_line, up_bar_fill &C.lxw_chart_fill, down_bar_line &C.lxw_chart_line, down_bar_fill &C.lxw_chart_fill) &C.lxw_chart
```
set_up_down_bars_format turn on up-down bars for the chart, with formatting

[[Return to contents]](#Contents)

## set_drop_lines
```v
fn (c &C.lxw_chart) set_drop_lines(line &C.lxw_chart_line) &C.lxw_chart
```
set_drop_lines turn on and format Drop Lines for a chart

[[Return to contents]](#Contents)

## set_high_low_lines
```v
fn (c &C.lxw_chart) set_high_low_lines(line &C.lxw_chart_line) &C.lxw_chart
```
set_high_low_lines turn on and format high-low Lines for a chart

[[Return to contents]](#Contents)

## set_series_overlap
```v
fn (c &C.lxw_chart) set_series_overlap(overlap i8) &C.lxw_chart
```
set_series_overlap set the `overlap`(-100 to 100) between series in a Bar/Column chart

[[Return to contents]](#Contents)

## set_series_gap
```v
fn (c &C.lxw_chart) set_series_gap(gap u16) &C.lxw_chart
```
set_series_gap set the `gap`(0 to 500) between series in a Bar/Column chart

[[Return to contents]](#Contents)

## show_blanks_as
```v
fn (c &C.lxw_chart) show_blanks_as(option Lxw_chart_blank) &C.lxw_chart
```
show_blanks_as set the option for displaying blank data in a chart

[[Return to contents]](#Contents)

## show_hidden_data
```v
fn (c &C.lxw_chart) show_hidden_data() &C.lxw_chart
```
show_hidden_data display data on charts from hidden rows or columns

[[Return to contents]](#Contents)

## set_rotation
```v
fn (c &C.lxw_chart) set_rotation(rotation u16) &C.lxw_chart
```
set_rotation set the Pie/Doughnut chart `rotation`(The angle of rotation)

[[Return to contents]](#Contents)

## set_hole_size
```v
fn (c &C.lxw_chart) set_hole_size(size u8) &C.lxw_chart
```
set_hole_size set the Doughnut chart hole `size`(The hole size as a percentage)

[[Return to contents]](#Contents)

## C.lxw_chart_axis
## set_name
```v
fn (a &C.lxw_chart_axis) set_name(name string) &C.lxw_chart_axis
```
set_name set the name caption of the an axis

[[Return to contents]](#Contents)

## set_name_range
```v
fn (a &C.lxw_chart_axis) set_name_range(sheetname string, row u32, col u16) &C.lxw_chart_axis
```
set_name_range set a chart axis name formula using row and column values

[[Return to contents]](#Contents)

## set_name_font
```v
fn (a &C.lxw_chart_axis) set_name_font(font &C.lxw_chart_font) &C.lxw_chart_axis
```
set_name_font set the font properties for a chart axis name

[[Return to contents]](#Contents)

## set_num_font
```v
fn (a &C.lxw_chart_axis) set_num_font(font &C.lxw_chart_font) &C.lxw_chart_axis
```
set_num_font set the font properties for the numbers of a chart axis

[[Return to contents]](#Contents)

## set_num_format
```v
fn (a &C.lxw_chart_axis) set_num_format(num_format string) &C.lxw_chart_axis
```
set_num_format set the number format for a chart axis

[[Return to contents]](#Contents)

## set_line
```v
fn (a &C.lxw_chart_axis) set_line(line &C.lxw_chart_line) &C.lxw_chart_axis
```
set_line set the line properties for a chart axis

[[Return to contents]](#Contents)

## set_fill
```v
fn (a &C.lxw_chart_axis) set_fill(fill &C.lxw_chart_fill) &C.lxw_chart_axis
```
set_fill set the fill properties for a chart axis

[[Return to contents]](#Contents)

## set_pattern
```v
fn (a &C.lxw_chart_axis) set_pattern(pattern &C.lxw_chart_pattern) &C.lxw_chart_axis
```
set_pattern set the pattern properties for a chart axis

[[Return to contents]](#Contents)

## set_reverse
```v
fn (a &C.lxw_chart_axis) set_reverse() &C.lxw_chart_axis
```
set_reverse reverse the order of the axis categories or values

[[Return to contents]](#Contents)

## set_crossing
```v
fn (a &C.lxw_chart_axis) set_crossing(value f64) &C.lxw_chart_axis
```
set_crossing set the position that the axis will cross the opposite axis

[[Return to contents]](#Contents)

## set_crossing_max
```v
fn (a &C.lxw_chart_axis) set_crossing_max() &C.lxw_chart_axis
```
set_crossing_max set the opposite axis crossing position as the axis maximum

[[Return to contents]](#Contents)

## set_crossing_min
```v
fn (a &C.lxw_chart_axis) set_crossing_min() &C.lxw_chart_axis
```
set_crossing_min set the opposite axis crossing position as the axis minimum

[[Return to contents]](#Contents)

## off
```v
fn (a &C.lxw_chart_axis) off() &C.lxw_chart_axis
```
off turn off/hide an axis

[[Return to contents]](#Contents)

## set_position
```v
fn (a &C.lxw_chart_axis) set_position(position Lxw_chart_axis_tick_position) &C.lxw_chart_axis
```
set_position position a category axis on or between the axis tick marks

[[Return to contents]](#Contents)

## set_label_position
```v
fn (a &C.lxw_chart_axis) set_label_position(position Lxw_chart_axis_label_position) &C.lxw_chart_axis
```
set_label_position position the axis labels

[[Return to contents]](#Contents)

## set_label_align
```v
fn (a &C.lxw_chart_axis) set_label_align(align Lxw_chart_axis_label_alignment) &C.lxw_chart_axis
```
set_label_align set the alignment of the axis labels

[[Return to contents]](#Contents)

## set_min
```v
fn (a &C.lxw_chart_axis) set_min(min f64) &C.lxw_chart_axis
```
set_min set the minimum value for a chart axis

[[Return to contents]](#Contents)

## set_max
```v
fn (a &C.lxw_chart_axis) set_max(max f64) &C.lxw_chart_axis
```
set_max set the maximum value for a chart axis

[[Return to contents]](#Contents)

## set_log_base
```v
fn (a &C.lxw_chart_axis) set_log_base(log_base u16) &C.lxw_chart_axis
```
set_log_base set the log base of the axis range

[[Return to contents]](#Contents)

## set_major_tick_mark
```v
fn (a &C.lxw_chart_axis) set_major_tick_mark(tick_mark_type Lxw_chart_tick_mark) &C.lxw_chart_axis
```
set_major_tick_mark set the major axis tick mark type

[[Return to contents]](#Contents)

## set_minor_tick_mark
```v
fn (a &C.lxw_chart_axis) set_minor_tick_mark(tick_mark_type Lxw_chart_tick_mark) &C.lxw_chart_axis
```
set_minor_tick_mark set the minor axis tick mark type

[[Return to contents]](#Contents)

## set_interval_unit
```v
fn (a &C.lxw_chart_axis) set_interval_unit(unit u16) &C.lxw_chart_axis
```
set_interval_unit set the interval between category values

[[Return to contents]](#Contents)

## set_interval_tick
```v
fn (a &C.lxw_chart_axis) set_interval_tick(unit u16) &C.lxw_chart_axis
```
set_interval_tick set the interval between category tick marks

[[Return to contents]](#Contents)

## set_major_unit
```v
fn (a &C.lxw_chart_axis) set_major_unit(unit f64) &C.lxw_chart_axis
```
set_major_unit set the increment of the major units in the axis

[[Return to contents]](#Contents)

## set_minor_unit
```v
fn (a &C.lxw_chart_axis) set_minor_unit(unit f64) &C.lxw_chart_axis
```
set_minor_unit set the increment of the minor units in the axis

[[Return to contents]](#Contents)

## set_display_units
```v
fn (a &C.lxw_chart_axis) set_display_units(units Lxw_chart_axis_display_unit) &C.lxw_chart_axis
```
set_display_units set the display units for a value axis

[[Return to contents]](#Contents)

## set_display_units_visible
```v
fn (a &C.lxw_chart_axis) set_display_units_visible(visible bool) &C.lxw_chart_axis
```
set_display_units_visible turn on/off the display units for a value axis

[[Return to contents]](#Contents)

## major_gridlines_set_visible
```v
fn (a &C.lxw_chart_axis) major_gridlines_set_visible(visible bool) &C.lxw_chart_axis
```
major_gridlines_set_visible turn on/off the major gridlines for an axis

[[Return to contents]](#Contents)

## minor_gridlines_set_visible
```v
fn (a &C.lxw_chart_axis) minor_gridlines_set_visible(visible bool) &C.lxw_chart_axis
```
minor_gridlines_set_visible turn on/off the minor gridlines for an axis

[[Return to contents]](#Contents)

## major_gridlines_set_line
```v
fn (a &C.lxw_chart_axis) major_gridlines_set_line(line &C.lxw_chart_line) &C.lxw_chart_axis
```
major_gridlines_set_line set the line properties for the chart axis major gridlines

[[Return to contents]](#Contents)

## minor_gridlines_set_line
```v
fn (a &C.lxw_chart_axis) minor_gridlines_set_line(line &C.lxw_chart_line) &C.lxw_chart_axis
```
minor_gridlines_set_line set the line properties for the chart axis minor gridlines

[[Return to contents]](#Contents)

## C.lxw_chart_series
## set_categories
```v
fn (s &C.lxw_chart_series) set_categories(sheetname string, first_row u32, first_col u16, last_row u32, last_col u16) &C.lxw_chart_series
```
set_categories set a series "categories" range using row and column values

[[Return to contents]](#Contents)

## set_values
```v
fn (s &C.lxw_chart_series) set_values(sheetname string, first_row u32, first_col u16, last_row u32, last_col u16) &C.lxw_chart_series
```
set_values set a series "values" range using row and column values

[[Return to contents]](#Contents)

## set_name
```v
fn (s &C.lxw_chart_series) set_name(name string) &C.lxw_chart_series
```
set_name set the name of a chart series range

[[Return to contents]](#Contents)

## set_name_range
```v
fn (s &C.lxw_chart_series) set_name_range(sheetname string, row u32, col u16) &C.lxw_chart_series
```
set_name_range set a series name formula using row and column values

[[Return to contents]](#Contents)

## set_line
```v
fn (s &C.lxw_chart_series) set_line(line &C.lxw_chart_line) &C.lxw_chart_series
```
set_line set the line properties for a chart series

[[Return to contents]](#Contents)

## set_fill
```v
fn (s &C.lxw_chart_series) set_fill(fill &C.lxw_chart_fill) &C.lxw_chart_series
```
set_fill set the fill properties for a chart series

[[Return to contents]](#Contents)

## set_invert_if_negative
```v
fn (s &C.lxw_chart_series) set_invert_if_negative() &C.lxw_chart_series
```
set_invert_if_negative invert the fill color for negative series values

[[Return to contents]](#Contents)

## set_pattern
```v
fn (s &C.lxw_chart_series) set_pattern(pattern &C.lxw_chart_pattern) &C.lxw_chart_series
```
set_pattern set the pattern properties for a chart series

[[Return to contents]](#Contents)

## set_marker_type
```v
fn (s &C.lxw_chart_series) set_marker_type(marker_type Lxw_chart_marker_type) &C.lxw_chart_series
```
set_marker_type set the data marker type for a series

[[Return to contents]](#Contents)

## set_marker_size
```v
fn (s &C.lxw_chart_series) set_marker_size(size u8) &C.lxw_chart_series
```
set_marker_size set the size of a data marker for a series

[[Return to contents]](#Contents)

## set_marker_line
```v
fn (s &C.lxw_chart_series) set_marker_line(line &C.lxw_chart_line) &C.lxw_chart_series
```
set_marker_line set the line properties for a chart series marker

[[Return to contents]](#Contents)

## set_marker_fill
```v
fn (s &C.lxw_chart_series) set_marker_fill(fill &C.lxw_chart_fill) &C.lxw_chart_series
```
set_marker_fill set the fill properties for a chart series marker

[[Return to contents]](#Contents)

## set_marker_pattern
```v
fn (s &C.lxw_chart_series) set_marker_pattern(pattern &C.lxw_chart_pattern) &C.lxw_chart_series
```
set_marker_pattern set the pattern properties for a chart series marker

[[Return to contents]](#Contents)

## set_points
```v
fn (s &C.lxw_chart_series) set_points(points &&C.lxw_chart_point) Lxw_error
```
set_points set the formatting for points in the series

[[Return to contents]](#Contents)

## set_smooth
```v
fn (s &C.lxw_chart_series) set_smooth(smooth bool) &C.lxw_chart_series
```
set_smooth smooth a line or scatter chart series

[[Return to contents]](#Contents)

## set_labels
```v
fn (s &C.lxw_chart_series) set_labels() &C.lxw_chart_series
```
set_labels add data labels to a chart series

[[Return to contents]](#Contents)

## set_labels_options
```v
fn (s &C.lxw_chart_series) set_labels_options(show_name bool, show_category bool, show_value bool) &C.lxw_chart_series
```
set_labels_options set the display options for the labels of a data series

[[Return to contents]](#Contents)

## set_labels_custom
```v
fn (s &C.lxw_chart_series) set_labels_custom(data_labels &&C.lxw_chart_data_label) Lxw_error
```
set_labels_custom set the properties for data labels in a series

[[Return to contents]](#Contents)

## set_labels_separator
```v
fn (s &C.lxw_chart_series) set_labels_separator(separator Lxw_chart_label_separator) &C.lxw_chart_series
```
set the separator for the data label captions

[[Return to contents]](#Contents)

## set_labels_position
```v
fn (s &C.lxw_chart_series) set_labels_position(position Lxw_chart_label_position) &C.lxw_chart_series
```
set_labels_position set the data label position for a series

[[Return to contents]](#Contents)

## set_labels_leader_line
```v
fn (s &C.lxw_chart_series) set_labels_leader_line() &C.lxw_chart_series
```
set_labels_leader_line set leader lines for Pie and Doughnut charts

[[Return to contents]](#Contents)

## set_labels_legend
```v
fn (s &C.lxw_chart_series) set_labels_legend() &C.lxw_chart_series
```
set_labels_legend set the legend key for a data label in a chart series

[[Return to contents]](#Contents)

## set_labels_percentage
```v
fn (s &C.lxw_chart_series) set_labels_percentage() &C.lxw_chart_series
```
set_labels_percentage set the percentage for a Pie/Doughnut data point

[[Return to contents]](#Contents)

## set_labels_num_format
```v
fn (s &C.lxw_chart_series) set_labels_num_format(num_format string) &C.lxw_chart_series
```
set_labels_num_format set the number format for chart data labels in a series

[[Return to contents]](#Contents)

## set_labels_font
```v
fn (s &C.lxw_chart_series) set_labels_font(font &C.lxw_chart_font) &C.lxw_chart_series
```
set_labels_font set the font properties for chart data labels in a series

[[Return to contents]](#Contents)

## set_labels_line
```v
fn (s &C.lxw_chart_series) set_labels_line(line &C.lxw_chart_line) &C.lxw_chart_series
```
set_labels_line set the line properties for the data labels in a chart series

[[Return to contents]](#Contents)

## set_labels_fill
```v
fn (s &C.lxw_chart_series) set_labels_fill(fill &C.lxw_chart_fill) &C.lxw_chart_series
```
set_labels_fill set the fill properties for the data labels in a chart series

[[Return to contents]](#Contents)

## set_labels_pattern
```v
fn (s &C.lxw_chart_series) set_labels_pattern(pattern &C.lxw_chart_pattern) &C.lxw_chart_series
```
set_labels_pattern set the pattern properties for the data labels in a chart series

[[Return to contents]](#Contents)

## set_trendline
```v
fn (s &C.lxw_chart_series) set_trendline(trendline_type Lxw_chart_trendline_type, value u8) &C.lxw_chart_series
```
set_trendline turn on a trendline for a chart data series

[[Return to contents]](#Contents)

## set_trendline_forecast
```v
fn (s &C.lxw_chart_series) set_trendline_forecast(forward f64, backward f64) &C.lxw_chart_series
```
set_trendline_forecast set the trendline forecast for a chart data series

[[Return to contents]](#Contents)

## set_trendline_equation
```v
fn (s &C.lxw_chart_series) set_trendline_equation() &C.lxw_chart_series
```
set_trendline_equation display the equation of a trendline for a chart data series

[[Return to contents]](#Contents)

## set_trendline_r_squared
```v
fn (s &C.lxw_chart_series) set_trendline_r_squared() &C.lxw_chart_series
```
set_trendline_r_squared display the R squared value of a trendline for a chart data series

[[Return to contents]](#Contents)

## set_trendline_intercept
```v
fn (s &C.lxw_chart_series) set_trendline_intercept(intercept f64) &C.lxw_chart_series
```
set_trendline_intercept set the trendline Y-axis intercept for a chart data series

[[Return to contents]](#Contents)

## set_trendline_name
```v
fn (s &C.lxw_chart_series) set_trendline_name(name string) &C.lxw_chart_series
```
set_trendline_name set the trendline name for a chart data series

[[Return to contents]](#Contents)

## set_trendline_line
```v
fn (s &C.lxw_chart_series) set_trendline_line(line &C.lxw_chart_line) &C.lxw_chart_series
```
set_trendline_line set the trendline line properties for a chart data series

[[Return to contents]](#Contents)

## get_error_bars
```v
fn (s &C.lxw_chart_series) get_error_bars(axis_type Lxw_chart_error_bar_axis) &C.lxw_series_error_bars
```
get_error_bars get a pointer to X or Y error bars from a chart series

[[Return to contents]](#Contents)

## C.lxw_chartsheet
## set_chart
```v
fn (c &C.lxw_chartsheet) set_chart(chart &C.lxw_chart) Lxw_error
```
set_chart insert a chart object into a chartsheet

[[Return to contents]](#Contents)

## activate
```v
fn (c &C.lxw_chartsheet) activate() &C.lxw_chartsheet
```
activate make a chartsheet the active, i.e., visible chartsheet

[[Return to contents]](#Contents)

## @select
```v
fn (c &C.lxw_chartsheet) @select() &C.lxw_chartsheet
```
@select set a chartsheet tab as selected

[[Return to contents]](#Contents)

## hide
```v
fn (c &C.lxw_chartsheet) hide() &C.lxw_chartsheet
```
hide hide the current chartsheet

[[Return to contents]](#Contents)

## set_first_sheet
```v
fn (c &C.lxw_chartsheet) set_first_sheet() &C.lxw_chartsheet
```
set_first_sheet set current chartsheet as the first visible sheet tab

[[Return to contents]](#Contents)

## set_tab_color
```v
fn (c &C.lxw_chartsheet) set_tab_color(color u32) &C.lxw_chartsheet
```
set_tab_color set the color of the chartsheet tab

[[Return to contents]](#Contents)

## protect
```v
fn (c &C.lxw_chartsheet) protect(password string, options &C.lxw_protection) &C.lxw_chartsheet
```
protect protect elements of a chartsheet from modification

[[Return to contents]](#Contents)

## set_zoom
```v
fn (c &C.lxw_chartsheet) set_zoom(scale u16) &C.lxw_chartsheet
```
set_zoom set the chartsheet zoom factor, `scale`(10 to 400)

[[Return to contents]](#Contents)

## set_landscape
```v
fn (c &C.lxw_chartsheet) set_landscape() &C.lxw_chartsheet
```
set_landscape set the page orientation as landscape

[[Return to contents]](#Contents)

## set_portrait
```v
fn (c &C.lxw_chartsheet) set_portrait() &C.lxw_chartsheet
```
set_portrait set the page orientation as portrait

[[Return to contents]](#Contents)

## set_paper
```v
fn (c &C.lxw_chartsheet) set_paper(paper_type Paper_format_type) &C.lxw_chartsheet
```
set_paper set the paper type for printing

[[Return to contents]](#Contents)

## set_margins
```v
fn (c &C.lxw_chartsheet) set_margins(left f64, right f64, top f64, bottom f64) &C.lxw_chartsheet
```
set_margins set the chartsheet margins for the printed page

[[Return to contents]](#Contents)

## set_header
```v
fn (c &C.lxw_chartsheet) set_header(str string) Lxw_error
```
set_header set the printed page header caption

[[Return to contents]](#Contents)

## set_footer
```v
fn (c &C.lxw_chartsheet) set_footer(str string) Lxw_error
```
set_footer set the printed page footer caption

[[Return to contents]](#Contents)

## set_header_opt
```v
fn (c &C.lxw_chartsheet) set_header_opt(str string, options &C.lxw_header_footer_options) Lxw_error
```
set_header_opt set the printed page header caption with additional options

[[Return to contents]](#Contents)

## set_footer_opt
```v
fn (c &C.lxw_chartsheet) set_footer_opt(str string, options &C.lxw_header_footer_options) Lxw_error
```
set_footer_opt set the printed page footer caption with additional options

[[Return to contents]](#Contents)

## C.lxw_format
## set_font_name
```v
fn (f &C.lxw_format) set_font_name(font_name string) &C.lxw_format
```
set_font_name set the font used in the cell

[[Return to contents]](#Contents)

## set_font_size
```v
fn (f &C.lxw_format) set_font_size(size f64) &C.lxw_format
```
set_font_size set the size of the font used in the cell

[[Return to contents]](#Contents)

## set_font_color
```v
fn (f &C.lxw_format) set_font_color(color u32) &C.lxw_format
```
set_font_color set the color of the font used in the cell

[[Return to contents]](#Contents)

## set_bold
```v
fn (f &C.lxw_format) set_bold() &C.lxw_format
```
set_bold turn on bold for the format font

[[Return to contents]](#Contents)

## set_italic
```v
fn (f &C.lxw_format) set_italic() &C.lxw_format
```
set_italic turn on italic for the format font

[[Return to contents]](#Contents)

## set_underline
```v
fn (f &C.lxw_format) set_underline(style Lxw_format_underlines) &C.lxw_format
```
set_underline turn on underline for the format

[[Return to contents]](#Contents)

## set_font_strikeout
```v
fn (f &C.lxw_format) set_font_strikeout() &C.lxw_format
```
set_font_strikeout set the strikeout property of the font

[[Return to contents]](#Contents)

## set_font_script
```v
fn (f &C.lxw_format) set_font_script(style Lxw_format_scripts) &C.lxw_format
```
set_font_script set the superscript/subscript property of the font

[[Return to contents]](#Contents)

## set_num_format
```v
fn (f &C.lxw_format) set_num_format(num_format string) &C.lxw_format
```
set_num_format set the number format for a cell

[[Return to contents]](#Contents)

## set_num_format_index
```v
fn (f &C.lxw_format) set_num_format_index(index u8) &C.lxw_format
```
set_num_format_index set the Excel built-in number format for a cell

[[Return to contents]](#Contents)

## set_unlocked
```v
fn (f &C.lxw_format) set_unlocked() &C.lxw_format
```
set_unlocked set the cell unlocked state

[[Return to contents]](#Contents)

## set_hidden
```v
fn (f &C.lxw_format) set_hidden() &C.lxw_format
```
set_hidden hide formulas in a cell

[[Return to contents]](#Contents)

## set_align
```v
fn (f &C.lxw_format) set_align(alignment Lxw_format_alignments) &C.lxw_format
```
set_align set the alignment for data in the cell

[[Return to contents]](#Contents)

## set_text_wrap
```v
fn (f &C.lxw_format) set_text_wrap() &C.lxw_format
```
set_text_wrap wrap text in a cell

[[Return to contents]](#Contents)

## set_rotation
```v
fn (f &C.lxw_format) set_rotation(angle i16) &C.lxw_format
```
set_rotation set the rotation of the text in a cell

[[Return to contents]](#Contents)

## set_indent
```v
fn (f &C.lxw_format) set_indent(level u8) &C.lxw_format
```
set_indent set the cell text indentation level

[[Return to contents]](#Contents)

## set_shrink
```v
fn (f &C.lxw_format) set_shrink() &C.lxw_format
```
set_shrink turn on the text "shrink to fit" for a cell

[[Return to contents]](#Contents)

## set_pattern
```v
fn (f &C.lxw_format) set_pattern(index Lxw_format_patterns) &C.lxw_format
```
set_pattern set the background fill pattern for a cell

[[Return to contents]](#Contents)

## set_bg_color
```v
fn (f &C.lxw_format) set_bg_color(color u32) &C.lxw_format
```
set_bg_color set the pattern background color for a cell

[[Return to contents]](#Contents)

## set_fg_color
```v
fn (f &C.lxw_format) set_fg_color(color u32) &C.lxw_format
```
set_fg_color set the pattern foreground color for a cell

[[Return to contents]](#Contents)

## set_border
```v
fn (f &C.lxw_format) set_border(style Lxw_format_borders) &C.lxw_format
```
set_border set the cell border style

[[Return to contents]](#Contents)

## set_bottom
```v
fn (f &C.lxw_format) set_bottom(style Lxw_format_borders) &C.lxw_format
```
set_bottom set the cell bottom border style

[[Return to contents]](#Contents)

## set_top
```v
fn (f &C.lxw_format) set_top(style Lxw_format_borders) &C.lxw_format
```
set_top set the cell top border style

[[Return to contents]](#Contents)

## set_left
```v
fn (f &C.lxw_format) set_left(style Lxw_format_borders) &C.lxw_format
```
set_left set the cell left border style

[[Return to contents]](#Contents)

## set_right
```v
fn (f &C.lxw_format) set_right(style Lxw_format_borders) &C.lxw_format
```
set_right set the cell right border style

[[Return to contents]](#Contents)

## set_border_color
```v
fn (f &C.lxw_format) set_border_color(color u32) &C.lxw_format
```
set_border_color set the color of the cell border

[[Return to contents]](#Contents)

## set_bottom_color
```v
fn (f &C.lxw_format) set_bottom_color(color u32) &C.lxw_format
```
set_bottom_color set the color of the bottom cell border

[[Return to contents]](#Contents)

## set_top_color
```v
fn (f &C.lxw_format) set_top_color(color u32) &C.lxw_format
```
set_top_color set the color of the top cell border

[[Return to contents]](#Contents)

## set_left_color
```v
fn (f &C.lxw_format) set_left_color(color u32) &C.lxw_format
```
set_left_color set the color of the left cell border

[[Return to contents]](#Contents)

## set_right_color
```v
fn (f &C.lxw_format) set_right_color(color u32) &C.lxw_format
```
set_right_color set the color of the right cell border

[[Return to contents]](#Contents)

## set_diag_type
```v
fn (f &C.lxw_format) set_diag_type(diagonal_type Lxw_format_diagonal_types) &C.lxw_format
```
set_diag_type set the diagonal cell border type

[[Return to contents]](#Contents)

## set_diag_border
```v
fn (f &C.lxw_format) set_diag_border(style Lxw_format_borders) &C.lxw_format
```
set_diag_border set the diagonal cell border style

[[Return to contents]](#Contents)

## set_diag_color
```v
fn (f &C.lxw_format) set_diag_color(color u32) &C.lxw_format
```
set_diag_color set the diagonal cell border color

[[Return to contents]](#Contents)

## set_quote_prefix
```v
fn (f &C.lxw_format) set_quote_prefix() &C.lxw_format
```
set_quote_prefix turn on quote prefix for the format

[[Return to contents]](#Contents)

## set_font_outline
```v
fn (f &C.lxw_format) set_font_outline() &C.lxw_format
```
set_font_outline set the font_outline property

[[Return to contents]](#Contents)

## set_font_shadow
```v
fn (f &C.lxw_format) set_font_shadow() &C.lxw_format
```
set_font_shadow set the font_shadow property

[[Return to contents]](#Contents)

## set_font_family
```v
fn (f &C.lxw_format) set_font_family(value u8) &C.lxw_format
```
set_font_family set the font_family property

[[Return to contents]](#Contents)

## set_font_charset
```v
fn (f &C.lxw_format) set_font_charset(value u8) &C.lxw_format
```
set_font_charset set the font_charset property

[[Return to contents]](#Contents)

## set_font_scheme
```v
fn (f &C.lxw_format) set_font_scheme(font_scheme string) &C.lxw_format
```
set_font_scheme set the font_scheme property

[[Return to contents]](#Contents)

## set_font_condense
```v
fn (f &C.lxw_format) set_font_condense() &C.lxw_format
```
set_font_condense set the font_condense property

[[Return to contents]](#Contents)

## set_font_extend
```v
fn (f &C.lxw_format) set_font_extend() &C.lxw_format
```
set_font_extend set the font_extend property

[[Return to contents]](#Contents)

## set_reading_order
```v
fn (f &C.lxw_format) set_reading_order(value u8) &C.lxw_format
```
set_reading_order set the reading_order property

[[Return to contents]](#Contents)

## set_theme
```v
fn (f &C.lxw_format) set_theme(value u8) &C.lxw_format
```
set_theme set the theme property

[[Return to contents]](#Contents)

## set_hyperlink
```v
fn (f &C.lxw_format) set_hyperlink() &C.lxw_format
```
set_hyperlink set the theme property

[[Return to contents]](#Contents)

## set_color_indexed
```v
fn (f &C.lxw_format) set_color_indexed(value u8) &C.lxw_format
```
set_color_indexed set the color_indexed property

[[Return to contents]](#Contents)

## set_font_only
```v
fn (f &C.lxw_format) set_font_only() &C.lxw_format
```
set_font_only set the font_only property

[[Return to contents]](#Contents)

## C.lxw_series_error_bars
## set
```v
fn (e &C.lxw_series_error_bars) set(error_bar_type Lxw_chart_error_bar_type, value f64) &C.lxw_series_error_bars
```
set set the X or Y error bars for a chart series

[[Return to contents]](#Contents)

## direction
```v
fn (e &C.lxw_series_error_bars) direction(direction Lxw_chart_error_bar_direction) &C.lxw_series_error_bars
```
direction set the direction (up, down or both) of the error bars for a chart series

[[Return to contents]](#Contents)

## endcap
```v
fn (e &C.lxw_series_error_bars) endcap(endcap Lxw_chart_error_bar_cap) &C.lxw_series_error_bars
```
endcap set the end cap type for the error bars of a chart series

[[Return to contents]](#Contents)

## line
```v
fn (e &C.lxw_series_error_bars) line(line &C.lxw_chart_line) &C.lxw_series_error_bars
```
line set the line properties for a chart series error bars

[[Return to contents]](#Contents)

## C.lxw_workbook
## add_worksheet
```v
fn (b &C.lxw_workbook) add_worksheet(sheetname string) &C.lxw_worksheet
```
add_worksheet add a new worksheet to a workbook

[[Return to contents]](#Contents)

## add_chartsheet
```v
fn (b &C.lxw_workbook) add_chartsheet(sheetname string) &C.lxw_chartsheet
```
add_chartsheet add a new chartsheet to a workbook

[[Return to contents]](#Contents)

## add_format
```v
fn (b &C.lxw_workbook) add_format() &C.lxw_format
```
add_format create a new Format object to formats cells in worksheets

[[Return to contents]](#Contents)

## add_chart
```v
fn (b &C.lxw_workbook) add_chart(chart_type Lxw_chart_type) &C.lxw_chart
```
add_chart create a new chart to be added to a worksheet

[[Return to contents]](#Contents)

## close
```v
fn (b &C.lxw_workbook) close() Lxw_error
```
close close the Workbook object and write the XLSX file

[[Return to contents]](#Contents)

## set_properties
```v
fn (b &C.lxw_workbook) set_properties(properties &C.lxw_doc_properties) Lxw_error
```
set_properties set the document properties such as Title, Author etc

[[Return to contents]](#Contents)

## set_custom_property
```v
fn (b &C.lxw_workbook) set_custom_property(name string, value All_Data_Type) Lxw_error
```
set_custom_property set a custom document property

[[Return to contents]](#Contents)

## define_name
```v
fn (b &C.lxw_workbook) define_name(name string, formula string) Lxw_error
```
define_name create a defined name in the workbook to use as a variable

[[Return to contents]](#Contents)

## get_default_url_format
```v
fn (b &C.lxw_workbook) get_default_url_format() &C.lxw_format
```
get_default_url_format get the default URL format used with write_url()

[[Return to contents]](#Contents)

## get_worksheet_by_name
```v
fn (b &C.lxw_workbook) get_worksheet_by_name(name string) &C.lxw_worksheet
```
get_worksheet_by_name get a worksheet object from its name

[[Return to contents]](#Contents)

## get_chartsheet_by_name
```v
fn (b &C.lxw_workbook) get_chartsheet_by_name(name string) &C.lxw_chartsheet
```
get_chartsheet_by_name get a chartsheet object from its name

[[Return to contents]](#Contents)

## validate_sheet_name
```v
fn (b &C.lxw_workbook) validate_sheet_name(sheetname string) Lxw_error
```
validate_sheet_name validate a worksheet or chartsheet name

[[Return to contents]](#Contents)

## add_vba_project
```v
fn (b &C.lxw_workbook) add_vba_project(filename string) Lxw_error
```
add_vba_project add a vbaProject binary to the Excel workbook

[[Return to contents]](#Contents)

## add_signed_vba_project
```v
fn (b &C.lxw_workbook) add_signed_vba_project(vba_project string, signature string) Lxw_error
```
add_signed_vba_project add a vbaProject binary and a vbaProjectSignature binary to the Excel workbook

[[Return to contents]](#Contents)

## set_vba_name
```v
fn (b &C.lxw_workbook) set_vba_name(name string) Lxw_error
```
set_vba_name set the VBA name for the workbook

[[Return to contents]](#Contents)

## read_only_recommended
```v
fn (b &C.lxw_workbook) read_only_recommended() &C.lxw_workbook
```
read_only_recommended add a recommendation to open the file in "read-only" mode

[[Return to contents]](#Contents)

## unset_default_url_format
```v
fn (b &C.lxw_workbook) unset_default_url_format() &C.lxw_workbook
```
unset_default_url_format unset the default URL format

[[Return to contents]](#Contents)

## C.lxw_worksheet
## write_struct
```v
fn (s &C.lxw_worksheet) write_struct[T](data []T)
```
write_struct write a struct array to xlsx file Example 1: struct Expense { item string              @[row:0;title:'ITEM'] // one record per row,from row 0; title is 'ITEM' cost vxlsx.All_Data_Type @[title:'Cost']       // title is 'Cost' } Example 2: struct Expense { item string              @[col:0;title:'ITEM'] // one record per col,from col 0; title is 'ITEM' cost vxlsx.All_Data_Type                       // title is 'cost' }

[[Return to contents]](#Contents)

## write
```v
fn (s &C.lxw_worksheet) write(row u32, col u16, value All_Data_Type) Lxw_error
```
write write `value` to a worksheet cell at (`row`:`col`)

[[Return to contents]](#Contents)

## write_number
```v
fn (s &C.lxw_worksheet) write_number(row u32, col u16, number f64, format &C.lxw_format) Lxw_error
```
write_number write a number to a worksheet cell

[[Return to contents]](#Contents)

## write_string
```v
fn (s &C.lxw_worksheet) write_string(row u32, col u16, str string, format &C.lxw_format) Lxw_error
```
write_string write a string to a worksheet cell

[[Return to contents]](#Contents)

## write_formula
```v
fn (s &C.lxw_worksheet) write_formula(row u32, col u16, formula string, format &C.lxw_format) Lxw_error
```
write_formula write a formula to a worksheet cell

[[Return to contents]](#Contents)

## write_array_formula
```v
fn (s &C.lxw_worksheet) write_array_formula(first_row u32, first_col u16, last_row u32, last_col u16, formula string, format &C.lxw_format) Lxw_error
```
write_array_formula write an array formula to a worksheet cell

[[Return to contents]](#Contents)

## write_dynamic_array_formula
```v
fn (s &C.lxw_worksheet) write_dynamic_array_formula(first_row u32, first_col u16, last_row u32, last_col u16, formula string, format &C.lxw_format) Lxw_error
```
write_dynamic_array_formula write an Excel 365 dynamic array formula to a worksheet range

[[Return to contents]](#Contents)

## write_dynamic_formula
```v
fn (s &C.lxw_worksheet) write_dynamic_formula(row u32, col u16, formula string, format &C.lxw_format) Lxw_error
```
write_dynamic_formula write an Excel 365 dynamic array formula to a worksheet cell

[[Return to contents]](#Contents)

## write_array_formula_num
```v
fn (s &C.lxw_worksheet) write_array_formula_num(first_row u32, first_col u16, last_row u32, last_col u16, formula string, format &C.lxw_format, result f64) Lxw_error
```
write_array_formula_num write an array formula with a numerical result to a cell in Excel

[[Return to contents]](#Contents)

## write_dynamic_array_formula_num
```v
fn (s &C.lxw_worksheet) write_dynamic_array_formula_num(first_row u32, first_col u16, last_row u32, last_col u16, formula string, format &C.lxw_format, result f64) Lxw_error
```
write_dynamic_array_formula_num write a dynamic array formula with a numerical result to a cell in Excel

[[Return to contents]](#Contents)

## write_dynamic_formula_num
```v
fn (s &C.lxw_worksheet) write_dynamic_formula_num(row u32, col u16, formula string, format &C.lxw_format, result f64) Lxw_error
```
write_dynamic_formula_num write a single cell dynamic array formula with a numerical result to a cell

[[Return to contents]](#Contents)

## write_datetime
```v
fn (s &C.lxw_worksheet) write_datetime(row u32, col u16, datetime &C.lxw_datetime, format &C.lxw_format) Lxw_error
```
write_datetime write a date or time to a worksheet cell

[[Return to contents]](#Contents)

## write_unixtime
```v
fn (s &C.lxw_worksheet) write_unixtime(row u32, col u16, unixtime i64, format &C.lxw_format) Lxw_error
```
write_unixtime write a Unix datetime to a worksheet cell

[[Return to contents]](#Contents)

## write_url
```v
fn (s &C.lxw_worksheet) write_url(row u32, col u16, url string, format &C.lxw_format) Lxw_error
```
write_url write a hyperlink/url to an Excel file

[[Return to contents]](#Contents)

## write_url_opt
```v
fn (s &C.lxw_worksheet) write_url_opt(row_num u32, col_num u16, url string, format &C.lxw_format, string_ &char, tooltip &char) Lxw_error
```
write_url_opt write a hyperlink/url to an Excel file

[[Return to contents]](#Contents)

## write_boolean
```v
fn (s &C.lxw_worksheet) write_boolean(row u32, col u16, value bool, format &C.lxw_format) Lxw_error
```
write_boolean write a formatted boolean worksheet cell

[[Return to contents]](#Contents)

## write_blank
```v
fn (s &C.lxw_worksheet) write_blank(row u32, col u16, format &C.lxw_format) Lxw_error
```
write_blank write a formatted blank worksheet cell

[[Return to contents]](#Contents)

## write_formula_num
```v
fn (s &C.lxw_worksheet) write_formula_num(row u32, col u16, formula string, format &C.lxw_format, result f64) Lxw_error
```
write_formula_num write a formula to a worksheet cell with a user defined numeric result

[[Return to contents]](#Contents)

## write_formula_str
```v
fn (s &C.lxw_worksheet) write_formula_str(row u32, col u16, formula string, format &C.lxw_format, result &char) Lxw_error
```
write_formula_str write a formula to a worksheet cell with a user defined string result

[[Return to contents]](#Contents)

## write_rich_string
```v
fn (s &C.lxw_worksheet) write_rich_string(row u32, col u16, rich_string &&C.lxw_rich_string_tuple, format &C.lxw_format) Lxw_error
```
write_rich_string write a "Rich" multi-format string to a worksheet cell

[[Return to contents]](#Contents)

## write_comment
```v
fn (s &C.lxw_worksheet) write_comment(row u32, col u16, str string) Lxw_error
```
write_comment write a comment to a worksheet cell

[[Return to contents]](#Contents)

## write_comment_opt
```v
fn (s &C.lxw_worksheet) write_comment_opt(row u32, col u16, str string, options &C.lxw_comment_options) Lxw_error
```
write_comment_opt write a comment to a worksheet cell with options

[[Return to contents]](#Contents)

## set_row
```v
fn (s &C.lxw_worksheet) set_row(row u32, height f64, format &C.lxw_format) Lxw_error
```
set_row set the properties for a row of cells

[[Return to contents]](#Contents)

## set_row_opt
```v
fn (s &C.lxw_worksheet) set_row_opt(row u32, height f64, format &C.lxw_format, options &C.lxw_row_col_options) Lxw_error
```
set_row_opt set the properties for a row of cells

[[Return to contents]](#Contents)

## set_row_pixels
```v
fn (s &C.lxw_worksheet) set_row_pixels(row u32, pixels u32, format &C.lxw_format) Lxw_error
```
set_row_pixels set the properties for a row of cells, with the height in pixels

[[Return to contents]](#Contents)

## set_row_pixels_opt
```v
fn (s &C.lxw_worksheet) set_row_pixels_opt(row u32, pixels u32, format &C.lxw_format, options &C.lxw_row_col_options) Lxw_error
```
set_row_pixels_opt set the properties for a row of cells, with the height in pixels

[[Return to contents]](#Contents)

## set_column
```v
fn (s &C.lxw_worksheet) set_column(first_col u16, last_col u16, width f64, format &C.lxw_format) Lxw_error
```
set_column set the properties for one or more columns of cells

[[Return to contents]](#Contents)

## set_column_opt
```v
fn (s &C.lxw_worksheet) set_column_opt(first_col u16, last_col u16, width f64, format &C.lxw_format, options &C.lxw_row_col_options) Lxw_error
```
set_column_opt set the properties for one or more columns of cells with options

[[Return to contents]](#Contents)

## set_column_pixels
```v
fn (s &C.lxw_worksheet) set_column_pixels(first_col u16, last_col u16, pixels u32, format &C.lxw_format) Lxw_error
```
set_column_pixels set the properties for one or more columns of cells, with the width in pixels

[[Return to contents]](#Contents)

## set_column_pixels_opt
```v
fn (s &C.lxw_worksheet) set_column_pixels_opt(first_col u16, last_col u16, pixels u32, format &C.lxw_format, options &C.lxw_row_col_options) Lxw_error
```
set_column_pixels_opt set the properties for one or more columns of cells with options, with the width in pixels

[[Return to contents]](#Contents)

## insert_image
```v
fn (s &C.lxw_worksheet) insert_image(row u32, col u16, filename string) Lxw_error
```
insert_image insert an image in a worksheet cell

[[Return to contents]](#Contents)

## insert_image_opt
```v
fn (s &C.lxw_worksheet) insert_image_opt(row u32, col u16, filename string, options &C.lxw_image_options) Lxw_error
```
insert_image_opt insert an image in a worksheet cell, with options

[[Return to contents]](#Contents)

## insert_image_buffer
```v
fn (s &C.lxw_worksheet) insert_image_buffer(row u32, col u16, image_buffer &u8, image_size usize) Lxw_error
```
insert_image_buffer insert an image in a worksheet cell, from a memory buffer

[[Return to contents]](#Contents)

## insert_image_buffer_opt
```v
fn (s &C.lxw_worksheet) insert_image_buffer_opt(row u32, col u16, image_buffer &u8, image_size usize, options &C.lxw_image_options) Lxw_error
```
insert_image_buffer_opt insert an image in a worksheet cell, from a memory buffer

[[Return to contents]](#Contents)

## set_background
```v
fn (s &C.lxw_worksheet) set_background(filename string) Lxw_error
```
set_background set the background image for a worksheet

[[Return to contents]](#Contents)

## set_background_buffer
```v
fn (s &C.lxw_worksheet) set_background_buffer(image_buffer &u8, image_size usize) Lxw_error
```
set_background_buffer set the background image for a worksheet, from a buffer

[[Return to contents]](#Contents)

## insert_chart
```v
fn (s &C.lxw_worksheet) insert_chart(row u32, col u16, chart &C.lxw_chart) Lxw_error
```
insert_chart insert a chart object into a worksheet

[[Return to contents]](#Contents)

## insert_chart_opt
```v
fn (s &C.lxw_worksheet) insert_chart_opt(row u32, col u16, chart &C.lxw_chart, user_options &C.lxw_chart_options) Lxw_error
```
insert_chart_opt insert a chart object into a worksheet, with options

[[Return to contents]](#Contents)

## merge_range
```v
fn (s &C.lxw_worksheet) merge_range(first_row u32, first_col u16, last_row u32, last_col u16, str string, format &C.lxw_format) Lxw_error
```
merge_range merge a range of cells

[[Return to contents]](#Contents)

## autofilter
```v
fn (s &C.lxw_worksheet) autofilter(first_row u32, first_col u16, last_row u32, last_col u16) Lxw_error
```
autofilter set the autofilter area in the worksheet

[[Return to contents]](#Contents)

## filter_column
```v
fn (s &C.lxw_worksheet) filter_column(col u16, rule &C.lxw_filter_rule) Lxw_error
```
filter_column write a filter rule to an autofilter column

[[Return to contents]](#Contents)

## filter_column2
```v
fn (s &C.lxw_worksheet) filter_column2(col u16, rule1 &C.lxw_filter_rule, rule2 &C.lxw_filter_rule, and_or u8) Lxw_error
```
filter_column2 write two filter rules to an autofilter column

[[Return to contents]](#Contents)

## filter_list
```v
fn (s &C.lxw_worksheet) filter_list(col u16, list &&u8) Lxw_error
```
filter_list write multiple string filters to an autofilter column

[[Return to contents]](#Contents)

## data_validation_cell
```v
fn (s &C.lxw_worksheet) data_validation_cell(row u32, col u16, validation &C.lxw_data_validation) Lxw_error
```
data_validation_cell add a data validation to a cell

[[Return to contents]](#Contents)

## data_validation_range
```v
fn (s &C.lxw_worksheet) data_validation_range(first_row u32, first_col u16, last_row u32, last_col u16, validation &C.lxw_data_validation) Lxw_error
```
data_validation_range add a data validation to a range

[[Return to contents]](#Contents)

## conditional_format_cell
```v
fn (s &C.lxw_worksheet) conditional_format_cell(row u32, col u16, conditional_format &C.lxw_conditional_format) Lxw_error
```
conditional_format_cell add a conditional format to a worksheet cell

[[Return to contents]](#Contents)

## conditional_format_range
```v
fn (s &C.lxw_worksheet) conditional_format_range(first_row u32, first_col u16, last_row u32, last_col u16, conditional_format &C.lxw_conditional_format) Lxw_error
```
conditional_format_range add a conditional format to a worksheet range

[[Return to contents]](#Contents)

## insert_button
```v
fn (s &C.lxw_worksheet) insert_button(row u32, col u16, options &C.lxw_button_options) Lxw_error
```
insert_button insert a button object into a worksheet

[[Return to contents]](#Contents)

## add_table
```v
fn (s &C.lxw_worksheet) add_table(first_row u32, first_col u16, last_row u32, last_col u16, options &C.lxw_table_options) Lxw_error
```
add_table add an Excel table to a worksheet

[[Return to contents]](#Contents)

## activate
```v
fn (s &C.lxw_worksheet) activate() &C.lxw_worksheet
```
activate make a worksheet the active, i.e., visible worksheet

[[Return to contents]](#Contents)

## @select
```v
fn (s &C.lxw_worksheet) @select() &C.lxw_worksheet
```
@select set a worksheet tab as selected

[[Return to contents]](#Contents)

## hide
```v
fn (s &C.lxw_worksheet) hide() &C.lxw_worksheet
```
hide hide the current worksheet

[[Return to contents]](#Contents)

## set_first_sheet
```v
fn (s &C.lxw_worksheet) set_first_sheet() &C.lxw_worksheet
```
set_first_sheet set current worksheet as the first visible sheet tab

[[Return to contents]](#Contents)

## freeze_panes
```v
fn (s &C.lxw_worksheet) freeze_panes(row u32, col u16) &C.lxw_worksheet
```
freeze_panes split and freeze a worksheet into panes

[[Return to contents]](#Contents)

## split_panes
```v
fn (s &C.lxw_worksheet) split_panes(vertical f64, horizontal f64) &C.lxw_worksheet
```
split_panes split a worksheet into panes

[[Return to contents]](#Contents)

## freeze_panes_opt
```v
fn (s &C.lxw_worksheet) freeze_panes_opt(first_row u32, first_col u16, top_row u32, left_col u16, type_ u8) &C.lxw_worksheet
```
freeze_panes_opt set panes and mark them as frozen. With extra options

[[Return to contents]](#Contents)

## split_panes_opt
```v
fn (s &C.lxw_worksheet) split_panes_opt(vertical f64, horizontal f64, top_row u32, left_col u16) &C.lxw_worksheet
```
split_panes_opt set panes and mark them as split.With extra options

[[Return to contents]](#Contents)

## set_selection
```v
fn (s &C.lxw_worksheet) set_selection(first_row u32, first_col u16, last_row u32, last_col u16) &C.lxw_worksheet
```
set_selection set the selected cell or cells in a worksheet

[[Return to contents]](#Contents)

## set_top_left_cell
```v
fn (s &C.lxw_worksheet) set_top_left_cell(row u32, col u16) &C.lxw_worksheet
```
set_top_left_cell set the first visible cell at the top left of a worksheet

[[Return to contents]](#Contents)

## set_landscape
```v
fn (s &C.lxw_worksheet) set_landscape() &C.lxw_worksheet
```
set_landscape set the page orientation as landscape

[[Return to contents]](#Contents)

## set_portrait
```v
fn (s &C.lxw_worksheet) set_portrait() &C.lxw_worksheet
```
set_portrait set the page orientation as portrait

[[Return to contents]](#Contents)

## set_page_view
```v
fn (s &C.lxw_worksheet) set_page_view() &C.lxw_worksheet
```
set_page_view set the page layout to page view mode

[[Return to contents]](#Contents)

## set_paper
```v
fn (s &C.lxw_worksheet) set_paper(paper_type Paper_format_type) &C.lxw_worksheet
```
set_paper set the paper type for printing

[[Return to contents]](#Contents)

## set_margins
```v
fn (s &C.lxw_worksheet) set_margins(left f64, right f64, top f64, bottom f64) &C.lxw_worksheet
```
set_margins set the worksheet margins for the printed page

[[Return to contents]](#Contents)

## set_header
```v
fn (s &C.lxw_worksheet) set_header(str string) Lxw_error
```
set_header set the printed page header caption

[[Return to contents]](#Contents)

## set_footer
```v
fn (s &C.lxw_worksheet) set_footer(str string) Lxw_error
```
set_footer set the printed page footer caption

[[Return to contents]](#Contents)

## set_header_opt
```v
fn (s &C.lxw_worksheet) set_header_opt(str string, options &C.lxw_header_footer_options) Lxw_error
```
set_header_opt set the printed page header caption with additional options

[[Return to contents]](#Contents)

## set_footer_opt
```v
fn (s &C.lxw_worksheet) set_footer_opt(str string, options &C.lxw_header_footer_options) Lxw_error
```
set_footer_opt set the printed page footer caption with additional options

[[Return to contents]](#Contents)

## set_h_pagebreaks
```v
fn (s &C.lxw_worksheet) set_h_pagebreaks(breaks &u32) Lxw_error
```
set_h_pagebreaks set the horizontal page breaks on a worksheet

[[Return to contents]](#Contents)

## set_v_pagebreaks
```v
fn (s &C.lxw_worksheet) set_v_pagebreaks(breaks &u16) Lxw_error
```
set_v_pagebreaks set the vertical page breaks on a worksheet

[[Return to contents]](#Contents)

## print_across
```v
fn (s &C.lxw_worksheet) print_across() &C.lxw_worksheet
```
print_across set the order in which pages are printed

[[Return to contents]](#Contents)

## set_zoom
```v
fn (s &C.lxw_worksheet) set_zoom(scale u16) &C.lxw_worksheet
```
set_zoom set the worksheet zoom factor

[[Return to contents]](#Contents)

## gridlines
```v
fn (s &C.lxw_worksheet) gridlines(option Lxw_gridlines) &C.lxw_worksheet
```
gridlines set the option to display or hide gridlines on the screen and the printed page

[[Return to contents]](#Contents)

## center_horizontally
```v
fn (s &C.lxw_worksheet) center_horizontally() &C.lxw_worksheet
```
center_horizontally center the printed page horizontally

[[Return to contents]](#Contents)

## center_vertically
```v
fn (s &C.lxw_worksheet) center_vertically() &C.lxw_worksheet
```
center_vertically center the printed page vertically

[[Return to contents]](#Contents)

## print_row_col_headers
```v
fn (s &C.lxw_worksheet) print_row_col_headers() &C.lxw_worksheet
```
print_row_col_headers set the option to print the row and column headers on the printed page

[[Return to contents]](#Contents)

## repeat_rows
```v
fn (s &C.lxw_worksheet) repeat_rows(first_row u32, last_row u32) Lxw_error
```
repeat_rows set the number of rows to repeat at the top of each printed page

[[Return to contents]](#Contents)

## repeat_columns
```v
fn (s &C.lxw_worksheet) repeat_columns(first_col u16, last_col u16) Lxw_error
```
repeat_columns set the number of columns to repeat at the top of each printed page

[[Return to contents]](#Contents)

## print_area
```v
fn (s &C.lxw_worksheet) print_area(first_row u32, first_col u16, last_row u32, last_col u16) Lxw_error
```
print_area set the print area for a worksheet

[[Return to contents]](#Contents)

## fit_to_pages
```v
fn (s &C.lxw_worksheet) fit_to_pages(width u16, height u16) &C.lxw_worksheet
```
fit_to_pages fit the printed area to a specific number of pages both vertically and horizontally

[[Return to contents]](#Contents)

## set_start_page
```v
fn (s &C.lxw_worksheet) set_start_page(start_page u16) &C.lxw_worksheet
```
set_start_page set the start/first page number when printing

[[Return to contents]](#Contents)

## set_print_scale
```v
fn (s &C.lxw_worksheet) set_print_scale(scale u16) &C.lxw_worksheet
```
set_print_scale set the scale factor for the printed page

[[Return to contents]](#Contents)

## print_black_and_white
```v
fn (s &C.lxw_worksheet) print_black_and_white() &C.lxw_worksheet
```
print_black_and_white set the worksheet to print in black and white

[[Return to contents]](#Contents)

## right_to_left
```v
fn (s &C.lxw_worksheet) right_to_left() &C.lxw_worksheet
```
right_to_left display the worksheet cells from right to left for some versions of Excel

[[Return to contents]](#Contents)

## hide_zero
```v
fn (s &C.lxw_worksheet) hide_zero() &C.lxw_worksheet
```
hide_zero hide zero values in worksheet cells

[[Return to contents]](#Contents)

## set_tab_color
```v
fn (s &C.lxw_worksheet) set_tab_color(color u32) &C.lxw_worksheet
```
set_tab_color set the color of the worksheet tab

[[Return to contents]](#Contents)

## protect
```v
fn (s &C.lxw_worksheet) protect(password string, options &C.lxw_protection) &C.lxw_worksheet
```
protect protect elements of a worksheet from modification

[[Return to contents]](#Contents)

## outline_settings
```v
fn (s &C.lxw_worksheet) outline_settings(visible u8, symbols_below u8, symbols_right u8, auto_style u8) &C.lxw_worksheet
```
outline_settings set the Outline and Grouping display properties

[[Return to contents]](#Contents)

## set_default_row
```v
fn (s &C.lxw_worksheet) set_default_row(height f64, hide_unused_rows u8) &C.lxw_worksheet
```
set_default_row set the default row properties

[[Return to contents]](#Contents)

## set_vba_name
```v
fn (s &C.lxw_worksheet) set_vba_name(name string) Lxw_error
```
set_vba_name set the VBA name for the worksheet

[[Return to contents]](#Contents)

## show_comments
```v
fn (s &C.lxw_worksheet) show_comments() &C.lxw_worksheet
```
show_comments make all comments in the worksheet visible

[[Return to contents]](#Contents)

## set_comments_author
```v
fn (s &C.lxw_worksheet) set_comments_author(author string) &C.lxw_worksheet
```
set_comments_author set the default author of the cell comments

[[Return to contents]](#Contents)

## ignore_errors
```v
fn (s &C.lxw_worksheet) ignore_errors(type_ u8, range string) Lxw_error
```
ignore_errors ignore various Excel errors/warnings in a worksheet for user defined ranges

[[Return to contents]](#Contents)

## Cell_types
[[Return to contents]](#Contents)

## Lxw_chart_axis_display_unit
```v
enum Lxw_chart_axis_display_unit {
	chart_axis_units_none // Axis display units: None. The default
	chart_axis_units_hundreds // Axis display units: Hundreds
	chart_axis_units_thousands // Axis display units: Thousands
	chart_axis_units_ten_thousands // Axis display units: Ten thousands
	chart_axis_units_hundred_thousands // Axis display units: Hundred thousands
	chart_axis_units_millions // Axis display units: Millions
	chart_axis_units_ten_millions // Axis display units: Ten millions
	chart_axis_units_hundred_millions // Axis display units: Hundred millions
	chart_axis_units_billions // Axis display units: Billions
	chart_axis_units_trillions // Axis display units: Trillions
}
```
Lxw_chart_axis_display_unit display units for chart value axis

[[Return to contents]](#Contents)

## Lxw_chart_axis_label_alignment
```v
enum Lxw_chart_axis_label_alignment {
	chart_axis_label_align_center // Chart axis label alignment: center
	chart_axis_label_align_left // Chart axis label alignment: left
	chart_axis_label_align_right // Chart axis label alignment: right
}
```
Lxw_chart_axis_label_alignment axis label alignments

[[Return to contents]](#Contents)

## Lxw_chart_axis_label_position
```v
enum Lxw_chart_axis_label_position {
	chart_axis_label_position_next_to // Position the axis labels next to the axis. The default
	chart_axis_label_position_high // Position the axis labels at the top of the chart, for horizontal axes, or to the right for vertical axes
	chart_axis_label_position_low // Position the axis labels at the bottom of the chart, for horizontal axes, or to the left for vertical axes
	chart_axis_label_position_none // Turn off the the axis labels
}
```
Lxw_chart_axis_label_position axis label positions

[[Return to contents]](#Contents)

## Lxw_chart_axis_tick_position
```v
enum Lxw_chart_axis_tick_position {
	chart_axis_position_default
	chart_axis_position_on_tick // Position category axis on tick marks
	chart_axis_position_between // Position category axis between tick marks
}
```
Lxw_chart_axis_tick_position axis positions for category axes

[[Return to contents]](#Contents)

## Lxw_chart_axis_type
```v
enum Lxw_chart_axis_type {
	chart_axis_type_x // Chart X axis
	chart_axis_type_y // Chart Y axis
}
```
Lxw_chart_axis_type chart axis types

[[Return to contents]](#Contents)

## Lxw_chart_blank
```v
enum Lxw_chart_blank {
	chart_blanks_as_gap // Show empty chart cells as gaps in the data. The default
	chart_blanks_as_zero // Show empty chart cells as zeros
	chart_blanks_as_connected // Show empty chart cells as connected. Only for charts with lines
}
```
Lxw_chart_blank define how blank values are displayed in a chart

[[Return to contents]](#Contents)

## Lxw_chart_error_bar_axis
```v
enum Lxw_chart_error_bar_axis {
	chart_error_bar_axis_x // X axis error bar
	chart_error_bar_axis_y // Y axis error bar
}
```
Lxw_chart_error_bar_axis direction for a data series error bar

[[Return to contents]](#Contents)

## Lxw_chart_error_bar_cap
```v
enum Lxw_chart_error_bar_cap {
	chart_error_bar_end_cap // Flat end cap. The default
	chart_error_bar_no_cap // No end cap
}
```
Lxw_chart_error_bar_cap end cap styles for a data series error bar

[[Return to contents]](#Contents)

## Lxw_chart_error_bar_direction
```v
enum Lxw_chart_error_bar_direction {
	chart_error_bar_dir_both // Error bar extends in both directions. The default
	chart_error_bar_dir_plus // Error bar extends in positive direction
	chart_error_bar_dir_minus // Error bar extends in negative direction
}
```
Lxw_chart_error_bar_direction direction for a data series error bar

[[Return to contents]](#Contents)

## Lxw_chart_error_bar_type
```v
enum Lxw_chart_error_bar_type {
	chart_error_bar_type_std_error // Error bar type: Standard error
	chart_error_bar_type_fixed // Error bar type: Fixed value
	chart_error_bar_type_percentage // Error bar type: Percentage
	chart_error_bar_type_std_dev // Error bar type: Standard deviation(s)
}
```
Lxw_chart_error_bar_type type/amount of data series error bar

[[Return to contents]](#Contents)

## Lxw_chart_grouping
[[Return to contents]](#Contents)

## Lxw_chart_label_position
```v
enum Lxw_chart_label_position {
	chart_label_position_default // Series data label position: default position
	chart_label_position_center // Series data label position: center
	chart_label_position_right // Series data label position: right
	chart_label_position_left // Series data label position: left
	chart_label_position_above // Series data label position: above
	chart_label_position_below // Series data label position: below
	chart_label_position_inside_base // Series data label position: inside base
	chart_label_position_inside_end // Series data label position: inside end
	chart_label_position_outside_end // Series data label position: outside end
	chart_label_position_best_fit // Series data label position: best fit
}
```
Lxw_chart_label_position chart data label positions

[[Return to contents]](#Contents)

## Lxw_chart_label_separator
```v
enum Lxw_chart_label_separator {
	chart_label_separator_comma // Series data label separator: comma (the default)
	chart_label_separator_semicolon // Series data label separator: semicolon
	chart_label_separator_period // Series data label separator: period
	chart_label_separator_newline // Series data label separator: newline
	chart_label_separator_space // Series data label separator: space
}
```
Lxw_chart_label_separator chart data label separator

[[Return to contents]](#Contents)

## Lxw_chart_legend_position
```v
enum Lxw_chart_legend_position {
	chart_legend_none              = 0 // No chart legend
	chart_legend_right // Chart legend positioned at right side
	chart_legend_left // Chart legend positioned at left side
	chart_legend_top // Chart legend positioned at top
	chart_legend_bottom // Chart legend positioned at bottom
	chart_legend_top_right // Chart legend positioned at top right
	chart_legend_overlay_right // Chart legend overlaid at right side
	chart_legend_overlay_left // Chart legend overlaid at left side
	chart_legend_overlay_top_right // Chart legend overlaid at top right
}
```
Lxw_chart_legend_position chart legend positions

[[Return to contents]](#Contents)

## Lxw_chart_line_dash_type
```v
enum Lxw_chart_line_dash_type {
	chart_line_dash_solid               = 0 // Solid
	chart_line_dash_round_dot // Round Dot
	chart_line_dash_square_dot // Square Dot
	chart_line_dash_dash // Dash
	chart_line_dash_dash_dot // Dash Dot
	chart_line_dash_long_dash // Long Dash
	chart_line_dash_long_dash_dot // Long Dash Dot
	chart_line_dash_long_dash_dot_dot // Long Dash Dot Dot
	chart_line_dash_dot // These aren't available in the dialog but are used by Excel
	chart_line_dash_system_dash_dot // These aren't available in the dialog but are used by Excel
	chart_line_dash_system_dash_dot_dot // These aren't available in the dialog but are used by Excel
}
```
Lxw_chart_line_dash_type chart line dash types

[[Return to contents]](#Contents)

## Lxw_chart_marker_type
```v
enum Lxw_chart_marker_type {
	chart_marker_automatic // Automatic, series default, marker type
	chart_marker_none // No marker type
	chart_marker_square // Square marker type
	chart_marker_diamond // Diamond marker type
	chart_marker_triangle // Triangle marker type
	chart_marker_x // X shape marker type
	chart_marker_star // Star marker type
	chart_marker_short_dash // Short dash marker type
	chart_marker_long_dash // Long dash marker type
	chart_marker_circle // Circle marker type
	chart_marker_plus // Plus (+) marker type
}
```
Lxw_chart_marker_type chart marker types

[[Return to contents]](#Contents)

## Lxw_chart_pattern_type
```v
enum Lxw_chart_pattern_type {
	chart_pattern_none // None pattern
	chart_pattern_percent_5 // 5 Percent pattern
	chart_pattern_percent_10 // 10 Percent pattern
	chart_pattern_percent_20 // 20 Percent pattern
	chart_pattern_percent_25 // 25 Percent pattern
	chart_pattern_percent_30 // 30 Percent pattern
	chart_pattern_percent_40 // 40 Percent pattern
	chart_pattern_percent_50 // 50 Percent pattern
	chart_pattern_percent_60 // 60 Percent pattern
	chart_pattern_percent_70 // 70 Percent pattern
	chart_pattern_percent_75 // 75 Percent pattern
	chart_pattern_percent_80 // 80 Percent pattern
	chart_pattern_percent_90 // 90 Percent pattern
	chart_pattern_light_downward_diagonal // Light downward diagonal pattern
	chart_pattern_light_upward_diagonal // Light upward diagonal pattern
	chart_pattern_dark_downward_diagonal // Dark downward diagonal pattern
	chart_pattern_dark_upward_diagonal // Dark upward diagonal pattern
	chart_pattern_wide_downward_diagonal // Wide downward diagonal pattern
	chart_pattern_wide_upward_diagonal // Wide upward diagonal pattern
	chart_pattern_light_vertical // Light vertical pattern
	chart_pattern_light_horizontal // Light horizontal pattern
	chart_pattern_narrow_vertical // Narrow vertical pattern
	chart_pattern_narrow_horizontal // Narrow horizontal pattern
	chart_pattern_dark_vertical // Dark vertical pattern
	chart_pattern_dark_horizontal // Dark horizontal pattern
	chart_pattern_dashed_downward_diagonal // Dashed downward diagonal pattern
	chart_pattern_dashed_upward_diagonal // Dashed upward diagonal pattern
	chart_pattern_dashed_horizontal //  Dashed horizontal pattern
	chart_pattern_dashed_vertical // Dashed vertical pattern
	chart_pattern_small_confetti // Small confetti pattern
	chart_pattern_large_confetti // Large confetti pattern
	chart_pattern_zigzag // Zigzag pattern
	chart_pattern_wave // Wave pattern
	chart_pattern_diagonal_brick // Diagonal brick pattern
	chart_pattern_horizontal_brick // Horizontal brick pattern
	chart_pattern_weave // Weave pattern
	chart_pattern_plaid // Plaid pattern
	chart_pattern_divot // Divot pattern
	chart_pattern_dotted_grid // Dotted grid pattern
	chart_pattern_dotted_diamond // Dotted diamond pattern
	chart_pattern_shingle // Shingle pattern
	chart_pattern_trellis // Trellis pattern
	chart_pattern_sphere // Sphere pattern
	chart_pattern_small_grid // Small grid pattern
	chart_pattern_large_grid // Large grid pattern
	chart_pattern_small_check // Small check pattern
	chart_pattern_large_check // Large check pattern
	chart_pattern_outlined_diamond // Outlined diamond pattern
	chart_pattern_solid_diamond // Solid diamond pattern
}
```
Lxw_chart_pattern_type chart pattern types

[[Return to contents]](#Contents)

## Lxw_chart_subtype
[[Return to contents]](#Contents)

## Lxw_chart_tick_mark
```v
enum Lxw_chart_tick_mark {
	chart_axis_tick_mark_default // Default tick mark for the chart axis. Usually outside
	chart_axis_tick_mark_none // No tick mark for the axis
	chart_axis_tick_mark_inside // Tick mark inside the axis only
	chart_axis_tick_mark_outside // Tick mark outside the axis only
	chart_axis_tick_mark_crossing // Tick mark inside and outside the axis
}
```
Lxw_chart_tick_mark tick mark types for an axis

[[Return to contents]](#Contents)

## Lxw_chart_trendline_type
```v
enum Lxw_chart_trendline_type {
	chart_trendline_type_linear // Trendline type: Linear
	chart_trendline_type_log // Trendline type: Logarithm
	chart_trendline_type_poly // Trendline type: Polynomial
	chart_trendline_type_power // Trendline type: Power
	chart_trendline_type_exp // Trendline type: Exponential
	chart_trendline_type_average // Trendline type: Moving Average
}
```
Lxw_chart_trendline_type series trendline/regression types

[[Return to contents]](#Contents)

## Lxw_chart_type
```v
enum Lxw_chart_type {
	chart_none                          = 0 // None
	chart_area // Area chart
	chart_area_stacked // Area chart - stacked
	chart_area_stacked_percent // Area chart - percentage stacked
	chart_bar // Bar chart
	chart_bar_stacked // Bar chart - stacked
	chart_bar_stacked_percent // Bar chart - percentage stacked
	chart_column // Column chart
	chart_column_stacked // Column chart - stacked
	chart_column_stacked_percent // Column chart - percentage stacked
	chart_doughnut // Doughnut chart
	chart_line // Line chart
	chart_line_stacked // Line chart - stacked
	chart_line_stacked_percent // Line chart - percentage stacked
	chart_pie // Pie chart
	chart_scatter // Scatter chart
	chart_scatter_straight // Scatter chart - straight
	chart_scatter_straight_with_markers // Scatter chart - straight with markers
	chart_scatter_smooth // Scatter chart - smooth
	chart_scatter_smooth_with_markers // Scatter chart - smooth with markers
	chart_radar // Radar chart
	chart_radar_with_markers // Radar chart - with markers
	chart_radar_filled // Radar chart - filled
}
```
Lxw_chart_type available chart types

[[Return to contents]](#Contents)

## Lxw_comment_display_types
```v
enum Lxw_comment_display_types {
	lxw_comment_display_default // Default to the worksheet default which can be hidden or visible
	lxw_comment_display_hidden // Hide the cell comment. Usually the default
	lxw_comment_display_visible // Show the cell comment. Can also be set for the worksheet with the worksheet.show_comments() function
}
```
Lxw_comment_display_types set the display type for a cell comment. This is hidden by default but can be set to visible with the worksheet.show_comments() function

[[Return to contents]](#Contents)

## Lxw_conditional_bar_axis_position
```v
enum Lxw_conditional_bar_axis_position {
	conditional_bar_axis_automatic // Data bar axis position is set by Excel based on the context of the data displayed
	conditional_bar_axis_midpoint // Data bar axis position is set at the midpoint
	conditional_bar_axis_none // Data bar axis is turned off
}
```
Lxw_conditional_bar_axis_position conditional format data bar axis options

[[Return to contents]](#Contents)

## Lxw_conditional_criteria
```v
enum Lxw_conditional_criteria {
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
```
Lxw_conditional_criteria the criteria used in a conditional format

[[Return to contents]](#Contents)

## Lxw_conditional_format_bar_direction
```v
enum Lxw_conditional_format_bar_direction {
	conditional_bar_direction_context // Data bar direction is set by Excel based on the context of the data displayed
	conditional_bar_direction_right_to_left // Data bar direction is from right to left
	conditional_bar_direction_left_to_right // Data bar direction is from left to right
}
```
Lxw_conditional_format_bar_direction conditional format data bar directions

[[Return to contents]](#Contents)

## Lxw_conditional_format_rule_types
```v
enum Lxw_conditional_format_rule_types {
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
```
Lxw_conditional_format_rule_types conditional format rule types

[[Return to contents]](#Contents)

## Lxw_conditional_format_types
```v
enum Lxw_conditional_format_types {
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
```
Lxw_conditional_format_types type definitions for conditional formats

[[Return to contents]](#Contents)

## Lxw_conditional_icon_types
```v
enum Lxw_conditional_icon_types {
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
```
Lxw_conditional_icon_types icon types used in the #lxw_conditional_format icon_style field

[[Return to contents]](#Contents)

## Lxw_defined_colors
```v
enum Lxw_defined_colors {
	color_black   = 0x1000000
	color_blue    = 0x0000FF
	color_brown   = 0x800000
	color_cyan    = 0x00FFFF
	color_gray    = 0x808080
	color_green   = 0x008000
	color_lime    = 0x00FF00
	color_magenta = 0xE4007F
	color_navy    = 0x000080
	color_orange  = 0xFF6600
	color_pink    = 0xFF00FF
	color_purple  = 0x800080
	color_red     = 0xFF0000
	color_silver  = 0xC0C0C0
	color_white   = 0xFFFFFF
	color_yellow  = 0xFFFF00
}
```
Lxw_defined_colors Predefined values for common colors

[[Return to contents]](#Contents)

## Lxw_error
[[Return to contents]](#Contents)

## Lxw_filter_criteria
```v
enum Lxw_filter_criteria {
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
```
Lxw_filter_criteria the criteria used in autofilter rules

[[Return to contents]](#Contents)

## Lxw_filter_operator
```v
enum Lxw_filter_operator {
	filter_and // Logical "and" of 2 filter rules
	filter_or // Logical "or" of 2 filter rules
}
```
Lxw_filter_operator and/or operator when using 2 filter rules

[[Return to contents]](#Contents)

## Lxw_filter_type
```v
enum Lxw_filter_type {
	filter_type_none
	filter_type_single
	filter_type_and
	filter_type_or
	filter_type_string_list
}
```
Lxw_filter_type internal filter types

[[Return to contents]](#Contents)

## Lxw_format_alignments
```v
enum Lxw_format_alignments {
	align_none                 = 0 // No alignment. Cell will use Excel's default for the data type
	align_left // Left horizontal alignment
	align_center // Center horizontal alignment
	align_right // Right horizontal alignment
	align_fill // Cell fill horizontal alignment
	align_justify // Justify horizontal alignment
	align_center_across // Center Across horizontal alignment
	align_distributed // Left horizontal alignment
	align_vertical_top // Top vertical alignment
	align_vertical_bottom // Bottom vertical alignment
	align_vertical_center // Center vertical alignment
	align_vertical_justify // Justify vertical alignment
	align_vertical_distributed // Distributed vertical alignment
}
```
Lxw_format_alignments alignment values for format.set_align()

[[Return to contents]](#Contents)

## Lxw_format_borders
```v
enum Lxw_format_borders {
	border_none // No border
	border_thin // Thin border style
	border_medium // Medium border style
	border_dashed // Dashed border style
	border_dotted // Dotted border style
	border_thick // Thick border style
	border_double // Double border style
	border_hair // Hair border style
	border_medium_dashed // Medium dashed border style
	border_dash_dot // Dash-dot border style
	border_medium_dash_dot // Medium dash-dot border style
	border_dash_dot_dot // Dash-dot-dot border style
	border_medium_dash_dot_dot // Medium dash-dot-dot border style
	border_slant_dash_dot // Slant dash-dot border style
}
```
Lxw_format_borders cell border styles for use with format.set_border()

[[Return to contents]](#Contents)

## Lxw_format_diagonal_types
```v
enum Lxw_format_diagonal_types {
	diagonal_border_up      = 1 // Cell diagonal border from bottom left to top right
	diagonal_border_down // Cell diagonal border from top left to bottom right
	diagonal_border_up_down // Cell diagonal border in both directions
}
```
Lxw_format_diagonal_types Diagonal border types

[[Return to contents]](#Contents)

## Lxw_format_patterns
```v
enum Lxw_format_patterns {
	pattern_none             = 0 // Empty pattern
	pattern_solid // Solid pattern
	pattern_medium_gray // Medium gray pattern
	pattern_dark_gray // Dark gray pattern
	pattern_light_gray // Light gray pattern
	pattern_dark_horizontal // Dark horizontal line pattern
	pattern_dark_vertical // Dark vertical line pattern
	pattern_dark_down // Dark diagonal stripe pattern
	pattern_dark_up // Reverse dark diagonal stripe pattern
	pattern_dark_grid // Dark grid pattern
	pattern_dark_trellis // Dark trellis pattern
	pattern_light_horizontal // Light horizontal Line pattern
	pattern_light_vertical // Light vertical line pattern
	pattern_light_down // Light diagonal stripe pattern
	pattern_light_up // Reverse light diagonal stripe pattern
	pattern_light_grid // Light grid pattern
	pattern_light_trellis // Light trellis pattern
	pattern_gray_125 // 12.5% gray pattern
	pattern_gray_0625 // 6.25% gray pattern
}
```
Lxw_format_patterns pattern value for use with format.set_pattern()

[[Return to contents]](#Contents)

## Lxw_format_scripts
```v
enum Lxw_format_scripts {
	font_superscript = 1 // Superscript font
	font_subscript // Subscript font
}
```
Lxw_format_scripts superscript and subscript values for format.set_font_script()

[[Return to contents]](#Contents)

## Lxw_format_underlines
```v
enum Lxw_format_underlines {
	underline_none              = 0
	underline_single // Single underline
	underline_double // Double underline
	underline_single_accounting // Single accounting underline
	underline_double_accounting // Double accounting underline
}
```
Lxw_format_underlines underline values for format.set_underline()

[[Return to contents]](#Contents)

## Lxw_gridlines
```v
enum Lxw_gridlines {
	hide_all_gridlines    = 0 // Hide screen and print gridlines
	show_screen_gridlines // Show screen gridlines
	show_print_gridlines // Show print gridlines
	show_all_gridlines // Show screen and print gridlines
}
```
Lxw_gridlines gridline options using in worksheet.gridlines()

[[Return to contents]](#Contents)

## Lxw_ignore_errors
```v
enum Lxw_ignore_errors {
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
```
Lxw_ignore_errors options for ignoring worksheet errors/warnings. See worksheet.ignore_errors()

[[Return to contents]](#Contents)

## Lxw_image_position
[[Return to contents]](#Contents)

## Lxw_object_position
```v
enum Lxw_object_position {
	object_position_default // Default positioning for the object
	object_move_and_size // Move and size the worksheet object with the cells
	object_move_dont_size // Move but don't size the worksheet object with the cells
	object_dont_move_dont_size // Don't move or size the worksheet object with the cells
	object_move_and_size_after // Same as #LXW_OBJECT_MOVE_AND_SIZE except libxlsxwriter applies hidden cells after the object is inserted
}
```
Lxw_object_position options to control the positioning of worksheet objects such as images or charts

[[Return to contents]](#Contents)

## Lxw_table_style_type
```v
enum Lxw_table_style_type {
	table_style_type_default
	table_style_type_light // Light table style
	table_style_type_medium // Medium table style
	table_style_type_dark // Dark table style
}
```
Lxw_table_style_type the type of table style

[[Return to contents]](#Contents)

## Lxw_table_total_functions
```v
enum Lxw_table_total_functions {
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
```
Lxw_table_total_functions standard Excel functions for totals in tables

[[Return to contents]](#Contents)

## Lxw_validation_boolean
```v
enum Lxw_validation_boolean {
	validation_default
	validation_off // Turn a data validation property off
	validation_on // Turn a data validation property on. Data validation properties are generally on by default
}
```
Lxw_validation_boolean data validation property values

[[Return to contents]](#Contents)

## Lxw_validation_criteria
```v
enum Lxw_validation_criteria {
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
```
Lxw_validation_criteria data validation criteria uses to control the selection of data

[[Return to contents]](#Contents)

## Lxw_validation_error_types
```v
enum Lxw_validation_error_types {
	validation_error_type_stop // Show a "Stop" data validation pop-up message. This is the default
	validation_error_type_warning // Show an "Error" data validation pop-up message
	validation_error_type_information // Show an "Information" data validation pop-up message
}
```
Lxw_validation_error_types data validation error types for pop-up messages

[[Return to contents]](#Contents)

## Lxw_validation_types
```v
enum Lxw_validation_types {
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
```
Lxw_validation_types data validation types

[[Return to contents]](#Contents)

## Pane_types
[[Return to contents]](#Contents)

## Paper_format_type
[[Return to contents]](#Contents)

## C.lxw_button_options
[[Return to contents]](#Contents)

## C.lxw_chart_fill
[[Return to contents]](#Contents)

## C.lxw_chart_line
[[Return to contents]](#Contents)

## C.lxw_chart_options
[[Return to contents]](#Contents)

## C.lxw_chart_pattern
[[Return to contents]](#Contents)

## C.lxw_comment_options
[[Return to contents]](#Contents)

## C.lxw_doc_properties
[[Return to contents]](#Contents)

## C.lxw_header_footer_options
[[Return to contents]](#Contents)

## C.lxw_image_options
[[Return to contents]](#Contents)

## C.lxw_protection
[[Return to contents]](#Contents)

## C.lxw_row_col_options
[[Return to contents]](#Contents)

## C.lxw_table_options
[[Return to contents]](#Contents)

## C.lxw_workbook_options
[[Return to contents]](#Contents)

## WorkBook_Options
[[Return to contents]](#Contents)

#### Powered by vdoc. Generated on: 13 Feb 2024 12:15:14
