module vxlsx

@[typedef]
struct C.lxw_chart_axis {}

@[typedef]
pub struct C.lxw_chart_line {
pub mut:
	color        u32
	none_        u8
	width        f32
	dash_type    u8
	transparency u8
}

@[typedef]
struct C.lxw_chart_font {}

@[typedef]
pub struct C.lxw_chart_pattern {
pub mut:
	fg_color u32
	bg_color u32
	@type    u8
}

@[typedef]
struct C.lxw_series_error_bars {}

@[typedef]
struct C.lxw_chart_error_bar_axis {}

@[typedef]
struct C.lxw_chart_series {}

@[typedef]
struct C.lxw_chart_data_label {}

@[typedef]
struct C.lxw_chart_point {}

@[typedef]
pub struct C.lxw_chart_fill {
pub mut:
	color        u32
	@none        u8
	transparency u8
}

// Lxw_chart_type available chart types
pub enum Lxw_chart_type {
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

// Lxw_chart_legend_position chart legend positions
pub enum Lxw_chart_legend_position {
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

// Lxw_chart_line_dash_type chart line dash types
pub enum Lxw_chart_line_dash_type {
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

// Lxw_chart_marker_type chart marker types
pub enum Lxw_chart_marker_type {
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

// Lxw_chart_pattern_type chart pattern types
pub enum Lxw_chart_pattern_type {
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

// Lxw_chart_label_position chart data label positions
pub enum Lxw_chart_label_position {
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

// Lxw_chart_label_separator chart data label separator
pub enum Lxw_chart_label_separator {
	chart_label_separator_comma // Series data label separator: comma (the default)
	chart_label_separator_semicolon // Series data label separator: semicolon
	chart_label_separator_period // Series data label separator: period
	chart_label_separator_newline // Series data label separator: newline
	chart_label_separator_space // Series data label separator: space
}

// Lxw_chart_axis_type chart axis types
pub enum Lxw_chart_axis_type {
	chart_axis_type_x // Chart X axis
	chart_axis_type_y // Chart Y axis
}

pub enum Lxw_chart_subtype {
	chart_subtype_none            = 0
	chart_subtype_stacked
	chart_subtype_stacked_percent
}

pub enum Lxw_chart_grouping {
	grouping_clustered
	grouping_standard
	grouping_percentstacked
	grouping_stacked
}

// Lxw_chart_axis_tick_position axis positions for category axes
pub enum Lxw_chart_axis_tick_position {
	chart_axis_position_default
	chart_axis_position_on_tick // Position category axis on tick marks
	chart_axis_position_between // Position category axis between tick marks
}

// Lxw_chart_axis_label_position axis label positions
pub enum Lxw_chart_axis_label_position {
	chart_axis_label_position_next_to // Position the axis labels next to the axis. The default
	chart_axis_label_position_high // Position the axis labels at the top of the chart, for horizontal axes, or to the right for vertical axes
	chart_axis_label_position_low // Position the axis labels at the bottom of the chart, for horizontal axes, or to the left for vertical axes
	chart_axis_label_position_none // Turn off the the axis labels
}

// Lxw_chart_axis_label_alignment axis label alignments
pub enum Lxw_chart_axis_label_alignment {
	chart_axis_label_align_center // Chart axis label alignment: center
	chart_axis_label_align_left // Chart axis label alignment: left
	chart_axis_label_align_right // Chart axis label alignment: right
}

// Lxw_chart_axis_display_unit display units for chart value axis
pub enum Lxw_chart_axis_display_unit {
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

// Lxw_chart_tick_mark tick mark types for an axis
pub enum Lxw_chart_tick_mark {
	chart_axis_tick_mark_default // Default tick mark for the chart axis. Usually outside
	chart_axis_tick_mark_none // No tick mark for the axis
	chart_axis_tick_mark_inside // Tick mark inside the axis only
	chart_axis_tick_mark_outside // Tick mark outside the axis only
	chart_axis_tick_mark_crossing // Tick mark inside and outside the axis
}

// Lxw_chart_trendline_type series trendline/regression types
pub enum Lxw_chart_trendline_type {
	chart_trendline_type_linear // Trendline type: Linear
	chart_trendline_type_log // Trendline type: Logarithm
	chart_trendline_type_poly // Trendline type: Polynomial
	chart_trendline_type_power // Trendline type: Power
	chart_trendline_type_exp // Trendline type: Exponential
	chart_trendline_type_average // Trendline type: Moving Average
}

// Lxw_chart_error_bar_axis direction for a data series error bar
pub enum Lxw_chart_error_bar_axis {
	chart_error_bar_axis_x // X axis error bar
	chart_error_bar_axis_y // Y axis error bar
}

// Lxw_chart_error_bar_type type/amount of data series error bar
pub enum Lxw_chart_error_bar_type {
	chart_error_bar_type_std_error // Error bar type: Standard error
	chart_error_bar_type_fixed // Error bar type: Fixed value
	chart_error_bar_type_percentage // Error bar type: Percentage
	chart_error_bar_type_std_dev // Error bar type: Standard deviation(s)
}

// Lxw_chart_error_bar_direction direction for a data series error bar
pub enum Lxw_chart_error_bar_direction {
	chart_error_bar_dir_both // Error bar extends in both directions. The default
	chart_error_bar_dir_plus // Error bar extends in positive direction
	chart_error_bar_dir_minus // Error bar extends in negative direction
}

// Lxw_chart_error_bar_cap end cap styles for a data series error bar
pub enum Lxw_chart_error_bar_cap {
	chart_error_bar_end_cap // Flat end cap. The default
	chart_error_bar_no_cap // No end cap
}

// Lxw_chart_blank define how blank values are displayed in a chart
pub enum Lxw_chart_blank {
	chart_blanks_as_gap // Show empty chart cells as gaps in the data. The default
	chart_blanks_as_zero // Show empty chart cells as zeros
	chart_blanks_as_connected // Show empty chart cells as connected. Only for charts with lines
}

// ===========chart functions===========
fn C.chart_add_series(chart &C.lxw_chart, categories &char, values &char) &C.lxw_chart_series
fn C.chart_series_set_categories(series &C.lxw_chart_series, sheetname &char, first_row u32, first_col u16, last_row u32, last_col u16)
fn C.chart_series_set_values(series &C.lxw_chart_series, sheetname &char, first_row u32, first_col u16, last_row u32, last_col u16)
fn C.chart_series_set_name(series &C.lxw_chart_series, name &char)
fn C.chart_series_set_name_range(series &C.lxw_chart_series, sheetname &char, row u32, col u16)
fn C.chart_series_set_line(series &C.lxw_chart_series, line &C.lxw_chart_line)
fn C.chart_series_set_fill(series &C.lxw_chart_series, fill &C.lxw_chart_fill)
fn C.chart_series_set_invert_if_negative(series &C.lxw_chart_series)
fn C.chart_series_set_pattern(series &C.lxw_chart_series, pattern &C.lxw_chart_pattern)
fn C.chart_series_set_marker_type(series &C.lxw_chart_series, type_ u8)
fn C.chart_series_set_marker_size(series &C.lxw_chart_series, size u8)
fn C.chart_series_set_marker_line(series &C.lxw_chart_series, line &C.lxw_chart_line)
fn C.chart_series_set_marker_fill(series &C.lxw_chart_series, fill &C.lxw_chart_fill)
fn C.chart_series_set_marker_pattern(series &C.lxw_chart_series, pattern &C.lxw_chart_pattern)
fn C.chart_series_set_points(series &C.lxw_chart_series, points &&C.lxw_chart_point) Lxw_error
fn C.chart_series_set_smooth(series &C.lxw_chart_series, smooth u8)
fn C.chart_series_set_labels(series &C.lxw_chart_series)
fn C.chart_series_set_labels_options(series &C.lxw_chart_series, show_name u8, show_category u8, show_value u8)
fn C.chart_series_set_labels_custom(series &C.lxw_chart_series, data_labels &&C.lxw_chart_data_label) Lxw_error
fn C.chart_series_set_labels_separator(series &C.lxw_chart_series, separator u8)
fn C.chart_series_set_labels_position(series &C.lxw_chart_series, position u8)
fn C.chart_series_set_labels_leader_line(series &C.lxw_chart_series)
fn C.chart_series_set_labels_legend(series &C.lxw_chart_series)
fn C.chart_series_set_labels_percentage(series &C.lxw_chart_series)
fn C.chart_series_set_labels_num_format(series &C.lxw_chart_series, num_format &char)
fn C.chart_series_set_labels_font(series &C.lxw_chart_series, font &C.lxw_chart_font)
fn C.chart_series_set_labels_line(series &C.lxw_chart_series, line &C.lxw_chart_line)
fn C.chart_series_set_labels_fill(series &C.lxw_chart_series, fill &C.lxw_chart_fill)
fn C.chart_series_set_labels_pattern(series &C.lxw_chart_series, pattern &C.lxw_chart_pattern)
fn C.chart_series_set_trendline(series &C.lxw_chart_series, type_ u8, value u8)
fn C.chart_series_set_trendline_forecast(series &C.lxw_chart_series, forward f64, backward f64)
fn C.chart_series_set_trendline_equation(series &C.lxw_chart_series)
fn C.chart_series_set_trendline_r_squared(series &C.lxw_chart_series)
fn C.chart_series_set_trendline_intercept(series &C.lxw_chart_series, intercept f64)
fn C.chart_series_set_trendline_name(series &C.lxw_chart_series, name &char)
fn C.chart_series_set_trendline_line(series &C.lxw_chart_series, line &C.lxw_chart_line)
fn C.chart_series_get_error_bars(series &C.lxw_chart_series, axis_type u8) &C.lxw_series_error_bars
fn C.chart_series_set_error_bars(error_bars &C.lxw_series_error_bars, type_ u8, value f64)
fn C.chart_series_set_error_bars_direction(error_bars &C.lxw_series_error_bars, direction u8)
fn C.chart_series_set_error_bars_endcap(error_bars &C.lxw_series_error_bars, endcap u8)
fn C.chart_series_set_error_bars_line(error_bars &C.lxw_series_error_bars, line &C.lxw_chart_line)
fn C.chart_axis_get(chart &C.lxw_chart, axis_type u8) &C.lxw_chart_axis
fn C.chart_axis_set_name(axis &C.lxw_chart_axis, name &char)
fn C.chart_axis_set_name_range(axis &C.lxw_chart_axis, sheetname &char, row u32, col u16)
fn C.chart_axis_set_name_font(axis &C.lxw_chart_axis, font &C.lxw_chart_font)
fn C.chart_axis_set_num_font(axis &C.lxw_chart_axis, font &C.lxw_chart_font)
fn C.chart_axis_set_num_format(axis &C.lxw_chart_axis, num_format &char)
fn C.chart_axis_set_line(axis &C.lxw_chart_axis, line &C.lxw_chart_line)
fn C.chart_axis_set_fill(axis &C.lxw_chart_axis, fill &C.lxw_chart_fill)
fn C.chart_axis_set_pattern(axis &C.lxw_chart_axis, pattern &C.lxw_chart_pattern)
fn C.chart_axis_set_reverse(axis &C.lxw_chart_axis)
fn C.chart_axis_set_crossing(axis &C.lxw_chart_axis, value f64)
fn C.chart_axis_set_crossing_max(axis &C.lxw_chart_axis)
fn C.chart_axis_set_crossing_min(axis &C.lxw_chart_axis)
fn C.chart_axis_off(axis &C.lxw_chart_axis)
fn C.chart_axis_set_position(axis &C.lxw_chart_axis, position u8)
fn C.chart_axis_set_label_position(axis &C.lxw_chart_axis, position u8)
fn C.chart_axis_set_label_align(axis &C.lxw_chart_axis, align u8)
fn C.chart_axis_set_min(axis &C.lxw_chart_axis, min f64)
fn C.chart_axis_set_max(axis &C.lxw_chart_axis, max f64)
fn C.chart_axis_set_log_base(axis &C.lxw_chart_axis, log_base u16)
fn C.chart_axis_set_major_tick_mark(axis &C.lxw_chart_axis, type_ u8)
fn C.chart_axis_set_minor_tick_mark(axis &C.lxw_chart_axis, type_ u8)
fn C.chart_axis_set_interval_unit(axis &C.lxw_chart_axis, unit u16)
fn C.chart_axis_set_interval_tick(axis &C.lxw_chart_axis, unit u16)
fn C.chart_axis_set_major_unit(axis &C.lxw_chart_axis, unit f64)
fn C.chart_axis_set_minor_unit(axis &C.lxw_chart_axis, unit f64)
fn C.chart_axis_set_display_units(axis &C.lxw_chart_axis, units u8)
fn C.chart_axis_set_display_units_visible(axis &C.lxw_chart_axis, visible u8)
fn C.chart_axis_major_gridlines_set_visible(axis &C.lxw_chart_axis, visible u8)
fn C.chart_axis_minor_gridlines_set_visible(axis &C.lxw_chart_axis, visible u8)
fn C.chart_axis_major_gridlines_set_line(axis &C.lxw_chart_axis, line &C.lxw_chart_line)
fn C.chart_axis_minor_gridlines_set_line(axis &C.lxw_chart_axis, line &C.lxw_chart_line)
fn C.chart_title_set_name(chart &C.lxw_chart, name &char)
fn C.chart_title_set_name_range(chart &C.lxw_chart, sheetname &char, row u32, col u16)
fn C.chart_title_set_name_font(chart &C.lxw_chart, font &C.lxw_chart_font)
fn C.chart_title_off(chart &C.lxw_chart)
fn C.chart_legend_set_position(chart &C.lxw_chart, position u8)
fn C.chart_legend_set_font(chart &C.lxw_chart, font &C.lxw_chart_font)
fn C.chart_legend_delete_series(chart &C.lxw_chart, delete_series &i16) Lxw_error
fn C.chart_chartarea_set_line(chart &C.lxw_chart, line &C.lxw_chart_line)
fn C.chart_chartarea_set_fill(chart &C.lxw_chart, fill &C.lxw_chart_fill)
fn C.chart_chartarea_set_pattern(chart &C.lxw_chart, pattern &C.lxw_chart_pattern)
fn C.chart_plotarea_set_line(chart &C.lxw_chart, line &C.lxw_chart_line)
fn C.chart_plotarea_set_fill(chart &C.lxw_chart, fill &C.lxw_chart_fill)
fn C.chart_plotarea_set_pattern(chart &C.lxw_chart, pattern &C.lxw_chart_pattern)
fn C.chart_set_style(chart &C.lxw_chart, style_id u8)
fn C.chart_set_table(chart &C.lxw_chart)
fn C.chart_set_table_grid(chart &C.lxw_chart, horizontal u8, vertical u8, outline u8, legend_keys u8)
fn C.chart_set_table_font(chart &C.lxw_chart, font &C.lxw_chart_font)
fn C.chart_set_up_down_bars(chart &C.lxw_chart)
fn C.chart_set_up_down_bars_format(chart &C.lxw_chart, up_bar_line &C.lxw_chart_line, up_bar_fill &C.lxw_chart_fill, down_bar_line &C.lxw_chart_line, down_bar_fill &C.lxw_chart_fill)
fn C.chart_set_drop_lines(chart &C.lxw_chart, line &C.lxw_chart_line)
fn C.chart_set_high_low_lines(chart &C.lxw_chart, line &C.lxw_chart_line)
fn C.chart_set_series_overlap(chart &C.lxw_chart, overlap i8)
fn C.chart_set_series_gap(chart &C.lxw_chart, gap u16)
fn C.chart_show_blanks_as(chart &C.lxw_chart, option u8)
fn C.chart_show_hidden_data(chart &C.lxw_chart)
fn C.chart_set_rotation(chart &C.lxw_chart, rotation u16)
fn C.chart_set_hole_size(chart &C.lxw_chart, size u8)

// add_series add a data series to a chart
pub fn (c &C.lxw_chart) add_series(categories string, values string) &C.lxw_chart_series {
	return C.chart_add_series(c, categories.str, values.str)
}

// set_categories set a series "categories" range using row and column values
pub fn (s &C.lxw_chart_series) set_categories(sheetname string, first_row u32, first_col u16, last_row u32, last_col u16) &C.lxw_chart_series {
	C.chart_series_set_categories(s, sheetname.str, first_row, first_col, last_row, last_col)
	return unsafe { s }
}

// set_values set a series "values" range using row and column values
pub fn (s &C.lxw_chart_series) set_values(sheetname string, first_row u32, first_col u16, last_row u32, last_col u16) &C.lxw_chart_series {
	C.chart_series_set_values(s, sheetname.str, first_row, first_col, last_row, last_col)
	return unsafe { s }
}

// set_name set the name of a chart series range
pub fn (s &C.lxw_chart_series) set_name(name string) &C.lxw_chart_series {
	C.chart_series_set_name(s, name.str)
	return unsafe { s }
}

// set_name_range set a series name formula using row and column values
pub fn (s &C.lxw_chart_series) set_name_range(sheetname string, row u32, col u16) &C.lxw_chart_series {
	C.chart_series_set_name_range(s, sheetname.str, row, col)
	return unsafe { s }
}

// set_line set the line properties for a chart series
pub fn (s &C.lxw_chart_series) set_line(line &C.lxw_chart_line) &C.lxw_chart_series {
	C.chart_series_set_line(s, line)
	return unsafe { s }
}

// set_fill set the fill properties for a chart series
pub fn (s &C.lxw_chart_series) set_fill(fill &C.lxw_chart_fill) &C.lxw_chart_series {
	C.chart_series_set_fill(s, fill)
	return unsafe { s }
}

// set_invert_if_negative invert the fill color for negative series values
pub fn (s &C.lxw_chart_series) set_invert_if_negative() &C.lxw_chart_series {
	C.chart_series_set_invert_if_negative(s)
	return unsafe { s }
}

// set_pattern set the pattern properties for a chart series
pub fn (s &C.lxw_chart_series) set_pattern(pattern &C.lxw_chart_pattern) &C.lxw_chart_series {
	C.chart_series_set_pattern(s, pattern)
	return unsafe { s }
}

// set_marker_type set the data marker type for a series
pub fn (s &C.lxw_chart_series) set_marker_type(marker_type Lxw_chart_marker_type) &C.lxw_chart_series {
	C.chart_series_set_marker_type(s, u8(marker_type))
	return unsafe { s }
}

// set_marker_size set the size of a data marker for a series
pub fn (s &C.lxw_chart_series) set_marker_size(size u8) &C.lxw_chart_series {
	C.chart_series_set_marker_size(s, size)
	return unsafe { s }
}

// set_marker_line set the line properties for a chart series marker
pub fn (s &C.lxw_chart_series) set_marker_line(line &C.lxw_chart_line) &C.lxw_chart_series {
	C.chart_series_set_marker_line(s, line)
	return unsafe { s }
}

// set_marker_fill set the fill properties for a chart series marker
pub fn (s &C.lxw_chart_series) set_marker_fill(fill &C.lxw_chart_fill) &C.lxw_chart_series {
	C.chart_series_set_marker_fill(s, fill)
	return unsafe { s }
}

// set_marker_pattern set the pattern properties for a chart series marker
pub fn (s &C.lxw_chart_series) set_marker_pattern(pattern &C.lxw_chart_pattern) &C.lxw_chart_series {
	C.chart_series_set_marker_pattern(s, pattern)
	return unsafe { s }
}

// set_points set the formatting for points in the series
pub fn (s &C.lxw_chart_series) set_points(points &&C.lxw_chart_point) Lxw_error {
	return C.chart_series_set_points(s, points)
}

// set_smooth smooth a line or scatter chart series
pub fn (s &C.lxw_chart_series) set_smooth(smooth bool) &C.lxw_chart_series {
	C.chart_series_set_smooth(s, u8(smooth))
	return unsafe { s }
}

// set_labels add data labels to a chart series
pub fn (s &C.lxw_chart_series) set_labels() &C.lxw_chart_series {
	C.chart_series_set_labels(s)
	return unsafe { s }
}

// set_labels_options set the display options for the labels of a data series
pub fn (s &C.lxw_chart_series) set_labels_options(show_name bool, show_category bool, show_value bool) &C.lxw_chart_series {
	C.chart_series_set_labels_options(s, u8(show_name), u8(show_category), u8(show_value))
	return unsafe { s }
}

// set_labels_custom set the properties for data labels in a series
pub fn (s &C.lxw_chart_series) set_labels_custom(data_labels &&C.lxw_chart_data_label) Lxw_error {
	return C.chart_series_set_labels_custom(s, data_labels)
}

// set the separator for the data label captions
pub fn (s &C.lxw_chart_series) set_labels_separator(separator Lxw_chart_label_separator) &C.lxw_chart_series {
	C.chart_series_set_labels_separator(s, u8(separator))
	return unsafe { s }
}

// set_labels_position set the data label position for a series
pub fn (s &C.lxw_chart_series) set_labels_position(position Lxw_chart_label_position) &C.lxw_chart_series {
	C.chart_series_set_labels_position(s, u8(position))
	return unsafe { s }
}

// set_labels_leader_line set leader lines for Pie and Doughnut charts
pub fn (s &C.lxw_chart_series) set_labels_leader_line() &C.lxw_chart_series {
	C.chart_series_set_labels_leader_line(s)
	return unsafe { s }
}

// set_labels_legend set the legend key for a data label in a chart series
pub fn (s &C.lxw_chart_series) set_labels_legend() &C.lxw_chart_series {
	C.chart_series_set_labels_legend(s)
	return unsafe { s }
}

// set_labels_percentage set the percentage for a Pie/Doughnut data point
pub fn (s &C.lxw_chart_series) set_labels_percentage() &C.lxw_chart_series {
	C.chart_series_set_labels_percentage(s)
	return unsafe { s }
}

// set_labels_num_format set the number format for chart data labels in a series
pub fn (s &C.lxw_chart_series) set_labels_num_format(num_format string) &C.lxw_chart_series {
	C.chart_series_set_labels_num_format(s, num_format.str)
	return unsafe { s }
}

// set_labels_font set the font properties for chart data labels in a series
pub fn (s &C.lxw_chart_series) set_labels_font(font &C.lxw_chart_font) &C.lxw_chart_series {
	C.chart_series_set_labels_font(s, font)
	return unsafe { s }
}

// set_labels_line set the line properties for the data labels in a chart series
pub fn (s &C.lxw_chart_series) set_labels_line(line &C.lxw_chart_line) &C.lxw_chart_series {
	C.chart_series_set_labels_line(s, line)
	return unsafe { s }
}

// set_labels_fill set the fill properties for the data labels in a chart series
pub fn (s &C.lxw_chart_series) set_labels_fill(fill &C.lxw_chart_fill) &C.lxw_chart_series {
	C.chart_series_set_labels_fill(s, fill)
	return unsafe { s }
}

// set_labels_pattern set the pattern properties for the data labels in a chart series
pub fn (s &C.lxw_chart_series) set_labels_pattern(pattern &C.lxw_chart_pattern) &C.lxw_chart_series {
	C.chart_series_set_labels_pattern(s, pattern)
	return unsafe { s }
}

// set_trendline turn on a trendline for a chart data series
pub fn (s &C.lxw_chart_series) set_trendline(trendline_type Lxw_chart_trendline_type, value u8) &C.lxw_chart_series {
	C.chart_series_set_trendline(s, u8(trendline_type), value)
	return unsafe { s }
}

// set_trendline_forecast set the trendline forecast for a chart data series
pub fn (s &C.lxw_chart_series) set_trendline_forecast(forward f64, backward f64) &C.lxw_chart_series {
	C.chart_series_set_trendline_forecast(s, forward, backward)
	return unsafe { s }
}

// set_trendline_equation display the equation of a trendline for a chart data series
pub fn (s &C.lxw_chart_series) set_trendline_equation() &C.lxw_chart_series {
	C.chart_series_set_trendline_equation(s)
	return unsafe { s }
}

// set_trendline_r_squared display the R squared value of a trendline for a chart data series
pub fn (s &C.lxw_chart_series) set_trendline_r_squared() &C.lxw_chart_series {
	C.chart_series_set_trendline_r_squared(s)
	return unsafe { s }
}

// set_trendline_intercept set the trendline Y-axis intercept for a chart data series
pub fn (s &C.lxw_chart_series) set_trendline_intercept(intercept f64) &C.lxw_chart_series {
	C.chart_series_set_trendline_intercept(s, intercept)
	return unsafe { s }
}

// set_trendline_name set the trendline name for a chart data series
pub fn (s &C.lxw_chart_series) set_trendline_name(name string) &C.lxw_chart_series {
	C.chart_series_set_trendline_name(s, name.str)
	return unsafe { s }
}

// set_trendline_line set the trendline line properties for a chart data series
pub fn (s &C.lxw_chart_series) set_trendline_line(line &C.lxw_chart_line) &C.lxw_chart_series {
	C.chart_series_set_trendline_line(s, line)
	return unsafe { s }
}

// get_error_bars get a pointer to X or Y error bars from a chart series
pub fn (s &C.lxw_chart_series) get_error_bars(axis_type Lxw_chart_error_bar_axis) &C.lxw_series_error_bars {
	return C.chart_series_get_error_bars(s, u8(axis_type))
}

// set set the X or Y error bars for a chart series
pub fn (e &C.lxw_series_error_bars) set(error_bar_type Lxw_chart_error_bar_type, value f64) &C.lxw_series_error_bars {
	C.chart_series_set_error_bars(e, u8(error_bar_type), value)
	return unsafe { e }
}

// direction set the direction (up, down or both) of the error bars for a chart series
pub fn (e &C.lxw_series_error_bars) direction(direction Lxw_chart_error_bar_direction) &C.lxw_series_error_bars {
	C.chart_series_set_error_bars_direction(e, u8(direction))
	return unsafe { e }
}

// endcap set the end cap type for the error bars of a chart series
pub fn (e &C.lxw_series_error_bars) endcap(endcap Lxw_chart_error_bar_cap) &C.lxw_series_error_bars {
	C.chart_series_set_error_bars_endcap(e, u8(endcap))
	return unsafe { e }
}

// line set the line properties for a chart series error bars
pub fn (e &C.lxw_series_error_bars) line(line &C.lxw_chart_line) &C.lxw_series_error_bars {
	C.chart_series_set_error_bars_line(e, line)
	return unsafe { e }
}

// get get an axis pointer from a chart
pub fn (c &C.lxw_chart) get(axis_type Lxw_chart_axis_type) &C.lxw_chart_axis {
	return C.chart_axis_get(c, u8(axis_type))
}

// set_name set the name caption of the an axis
pub fn (a &C.lxw_chart_axis) set_name(name string) &C.lxw_chart_axis {
	C.chart_axis_set_name(a, name.str)
	return unsafe { a }
}

// set_name_range set a chart axis name formula using row and column values
pub fn (a &C.lxw_chart_axis) set_name_range(sheetname string, row u32, col u16) &C.lxw_chart_axis {
	C.chart_axis_set_name_range(a, sheetname.str, row, col)
	return unsafe { a }
}

// set_name_font set the font properties for a chart axis name
pub fn (a &C.lxw_chart_axis) set_name_font(font &C.lxw_chart_font) &C.lxw_chart_axis {
	C.chart_axis_set_name_font(a, font)
	return unsafe { a }
}

// set_num_font set the font properties for the numbers of a chart axis
pub fn (a &C.lxw_chart_axis) set_num_font(font &C.lxw_chart_font) &C.lxw_chart_axis {
	C.chart_axis_set_num_font(a, font)
	return unsafe { a }
}

// set_num_format set the number format for a chart axis
pub fn (a &C.lxw_chart_axis) set_num_format(num_format string) &C.lxw_chart_axis {
	C.chart_axis_set_num_format(a, num_format.str)
	return unsafe { a }
}

// set_line set the line properties for a chart axis
pub fn (a &C.lxw_chart_axis) set_line(line &C.lxw_chart_line) &C.lxw_chart_axis {
	C.chart_axis_set_line(a, line)
	return unsafe { a }
}

// set_fill set the fill properties for a chart axis
pub fn (a &C.lxw_chart_axis) set_fill(fill &C.lxw_chart_fill) &C.lxw_chart_axis {
	C.chart_axis_set_fill(a, fill)
	return unsafe { a }
}

// set_pattern set the pattern properties for a chart axis
pub fn (a &C.lxw_chart_axis) set_pattern(pattern &C.lxw_chart_pattern) &C.lxw_chart_axis {
	C.chart_axis_set_pattern(a, pattern)
	return unsafe { a }
}

// set_reverse reverse the order of the axis categories or values
pub fn (a &C.lxw_chart_axis) set_reverse() &C.lxw_chart_axis {
	C.chart_axis_set_reverse(a)
	return unsafe { a }
}

// set_crossing set the position that the axis will cross the opposite axis
pub fn (a &C.lxw_chart_axis) set_crossing(value f64) &C.lxw_chart_axis {
	C.chart_axis_set_crossing(a, value)
	return unsafe { a }
}

// set_crossing_max set the opposite axis crossing position as the axis maximum
pub fn (a &C.lxw_chart_axis) set_crossing_max() &C.lxw_chart_axis {
	C.chart_axis_set_crossing_max(a)
	return unsafe { a }
}

// set_crossing_min set the opposite axis crossing position as the axis minimum
pub fn (a &C.lxw_chart_axis) set_crossing_min() &C.lxw_chart_axis {
	C.chart_axis_set_crossing_min(a)
	return unsafe { a }
}

// off turn off/hide an axis
pub fn (a &C.lxw_chart_axis) off() &C.lxw_chart_axis {
	C.chart_axis_off(a)
	return unsafe { a }
}

// set_position position a category axis on or between the axis tick marks
pub fn (a &C.lxw_chart_axis) set_position(position Lxw_chart_axis_tick_position) &C.lxw_chart_axis {
	C.chart_axis_set_position(a, u8(position))
	return unsafe { a }
}

// set_label_position position the axis labels
pub fn (a &C.lxw_chart_axis) set_label_position(position Lxw_chart_axis_label_position) &C.lxw_chart_axis {
	C.chart_axis_set_label_position(a, u8(position))
	return unsafe { a }
}

// set_label_align set the alignment of the axis labels
pub fn (a &C.lxw_chart_axis) set_label_align(align Lxw_chart_axis_label_alignment) &C.lxw_chart_axis {
	C.chart_axis_set_label_align(a, u8(align))
	return unsafe { a }
}

// set_min set the minimum value for a chart axis
pub fn (a &C.lxw_chart_axis) set_min(min f64) &C.lxw_chart_axis {
	C.chart_axis_set_min(a, min)
	return unsafe { a }
}

// set_max set the maximum value for a chart axis
pub fn (a &C.lxw_chart_axis) set_max(max f64) &C.lxw_chart_axis {
	C.chart_axis_set_max(a, max)
	return unsafe { a }
}

// set_log_base set the log base of the axis range
pub fn (a &C.lxw_chart_axis) set_log_base(log_base u16) &C.lxw_chart_axis {
	C.chart_axis_set_log_base(a, log_base)
	return unsafe { a }
}

// set_major_tick_mark set the major axis tick mark type
pub fn (a &C.lxw_chart_axis) set_major_tick_mark(tick_mark_type Lxw_chart_tick_mark) &C.lxw_chart_axis {
	C.chart_axis_set_major_tick_mark(a, u8(tick_mark_type))
	return unsafe { a }
}

// set_minor_tick_mark set the minor axis tick mark type
pub fn (a &C.lxw_chart_axis) set_minor_tick_mark(tick_mark_type Lxw_chart_tick_mark) &C.lxw_chart_axis {
	C.chart_axis_set_minor_tick_mark(a, u8(tick_mark_type))
	return unsafe { a }
}

// set_interval_unit set the interval between category values
pub fn (a &C.lxw_chart_axis) set_interval_unit(unit u16) &C.lxw_chart_axis {
	C.chart_axis_set_interval_unit(a, unit)
	return unsafe { a }
}

// set_interval_tick set the interval between category tick marks
pub fn (a &C.lxw_chart_axis) set_interval_tick(unit u16) &C.lxw_chart_axis {
	C.chart_axis_set_interval_tick(a, unit)
	return unsafe { a }
}

// set_major_unit set the increment of the major units in the axis
pub fn (a &C.lxw_chart_axis) set_major_unit(unit f64) &C.lxw_chart_axis {
	C.chart_axis_set_major_unit(a, unit)
	return unsafe { a }
}

// set_minor_unit set the increment of the minor units in the axis
pub fn (a &C.lxw_chart_axis) set_minor_unit(unit f64) &C.lxw_chart_axis {
	C.chart_axis_set_minor_unit(a, unit)
	return unsafe { a }
}

// set_display_units set the display units for a value axis
pub fn (a &C.lxw_chart_axis) set_display_units(units Lxw_chart_axis_display_unit) &C.lxw_chart_axis {
	C.chart_axis_set_display_units(a, u8(units))
	return unsafe { a }
}

// set_display_units_visible turn on/off the display units for a value axis
pub fn (a &C.lxw_chart_axis) set_display_units_visible(visible bool) &C.lxw_chart_axis {
	C.chart_axis_set_display_units_visible(a, u8(visible))
	return unsafe { a }
}

// major_gridlines_set_visible turn on/off the major gridlines for an axis
pub fn (a &C.lxw_chart_axis) major_gridlines_set_visible(visible bool) &C.lxw_chart_axis {
	C.chart_axis_major_gridlines_set_visible(a, u8(visible))
	return unsafe { a }
}

// minor_gridlines_set_visible turn on/off the minor gridlines for an axis
pub fn (a &C.lxw_chart_axis) minor_gridlines_set_visible(visible bool) &C.lxw_chart_axis {
	C.chart_axis_minor_gridlines_set_visible(a, u8(visible))
	return unsafe { a }
}

// major_gridlines_set_line set the line properties for the chart axis major gridlines
pub fn (a &C.lxw_chart_axis) major_gridlines_set_line(line &C.lxw_chart_line) &C.lxw_chart_axis {
	C.chart_axis_major_gridlines_set_line(a, line)
	return unsafe { a }
}

// minor_gridlines_set_line set the line properties for the chart axis minor gridlines
pub fn (a &C.lxw_chart_axis) minor_gridlines_set_line(line &C.lxw_chart_line) &C.lxw_chart_axis {
	C.chart_axis_minor_gridlines_set_line(a, line)
	return unsafe { a }
}

// title_set_name set the title of the chart
pub fn (c &C.lxw_chart) title_set_name(name string) &C.lxw_chart {
	C.chart_title_set_name(c, name.str)
	return unsafe { c }
}

// title_set_name_range set a chart title formula using row and column values
pub fn (c &C.lxw_chart) title_set_name_range(sheetname string, row u32, col u16) &C.lxw_chart {
	C.chart_title_set_name_range(c, sheetname.str, row, col)
	return unsafe { c }
}

// title_set_name_font set the font properties for a chart title
pub fn (c &C.lxw_chart) title_set_name_font(font &C.lxw_chart_font) &C.lxw_chart {
	C.chart_title_set_name_font(c, font)
	return unsafe { c }
}

// title_off turn off an automatic chart title
pub fn (c &C.lxw_chart) title_off() &C.lxw_chart {
	C.chart_title_off(c)
	return unsafe { c }
}

// legend_set_position set the position of the chart legend
pub fn (c &C.lxw_chart) legend_set_position(position Lxw_chart_legend_position) &C.lxw_chart {
	C.chart_legend_set_position(c, u8(position))
	return unsafe { c }
}

// legend_set_font set the font properties for a chart legend
pub fn (c &C.lxw_chart) legend_set_font(font &C.lxw_chart_font) &C.lxw_chart {
	C.chart_legend_set_font(c, font)
	return unsafe { c }
}

// legend_delete_series remove one or more series from the the legend
pub fn (c &C.lxw_chart) legend_delete_series(delete_series &i16) Lxw_error {
	return C.chart_legend_delete_series(c, delete_series)
}

// chartarea_set_line set the line properties for a chartarea
pub fn (c &C.lxw_chart) chartarea_set_line(line &C.lxw_chart_line) &C.lxw_chart {
	C.chart_chartarea_set_line(c, line)
	return unsafe { c }
}

// chartarea_set_fil set the fill properties for a chartareal
pub fn (c &C.lxw_chart) chartarea_set_fill(fill &C.lxw_chart_fill) &C.lxw_chart {
	C.chart_chartarea_set_fill(c, fill)
	return unsafe { c }
}

// chartarea_set_pattern set the pattern properties for a chartarea
pub fn (c &C.lxw_chart) chartarea_set_pattern(pattern &C.lxw_chart_pattern) &C.lxw_chart {
	C.chart_chartarea_set_pattern(c, pattern)
	return unsafe { c }
}

// plotarea_set_line set the line properties for a plotarea
pub fn (c &C.lxw_chart) plotarea_set_line(line &C.lxw_chart_line) &C.lxw_chart {
	C.chart_plotarea_set_line(c, line)
	return unsafe { c }
}

// plotarea_set_fill set the fill properties for a plotarea
pub fn (c &C.lxw_chart) plotarea_set_fill(fill &C.lxw_chart_fill) &C.lxw_chart {
	C.chart_plotarea_set_fill(c, fill)
	return unsafe { c }
}

// plotarea_set_pattern set the pattern properties for a plotarea
pub fn (c &C.lxw_chart) plotarea_set_pattern(pattern &C.lxw_chart_pattern) &C.lxw_chart {
	C.chart_plotarea_set_pattern(c, pattern)
	return unsafe { c }
}

// set_style set the chart style type, `style_id` is an index representing the chart style, 1 - 48
pub fn (c &C.lxw_chart) set_style(style_id u8) &C.lxw_chart {
	C.chart_set_style(c, style_id)
	return unsafe { c }
}

// set_table turn on a data table below the horizontal axis
pub fn (c &C.lxw_chart) set_table() &C.lxw_chart {
	C.chart_set_table(c)
	return unsafe { c }
}

// set_table_grid turn on/off grid options for a chart data table
pub fn (c &C.lxw_chart) set_table_grid(horizontal bool, vertical bool, outline bool, legend_keys bool) &C.lxw_chart {
	C.chart_set_table_grid(c, u8(horizontal), u8(vertical), u8(outline), u8(legend_keys))
	return unsafe { c }
}

// set_table_font
pub fn (c &C.lxw_chart) set_table_font(font &C.lxw_chart_font) &C.lxw_chart {
	C.chart_set_table_font(c, font)
	return unsafe { c }
}

// set_up_down_bars turn on up-down bars for the chart
pub fn (c &C.lxw_chart) set_up_down_bars() &C.lxw_chart {
	C.chart_set_up_down_bars(c)
	return unsafe { c }
}

// set_up_down_bars_format turn on up-down bars for the chart, with formatting
pub fn (c &C.lxw_chart) set_up_down_bars_format(up_bar_line &C.lxw_chart_line, up_bar_fill &C.lxw_chart_fill, down_bar_line &C.lxw_chart_line, down_bar_fill &C.lxw_chart_fill) &C.lxw_chart {
	C.chart_set_up_down_bars_format(c, up_bar_line, up_bar_fill, down_bar_line, down_bar_fill)
	return unsafe { c }
}

// set_drop_lines turn on and format Drop Lines for a chart
pub fn (c &C.lxw_chart) set_drop_lines(line &C.lxw_chart_line) &C.lxw_chart {
	C.chart_set_drop_lines(c, line)
	return unsafe { c }
}

// set_high_low_lines turn on and format high-low Lines for a chart
pub fn (c &C.lxw_chart) set_high_low_lines(line &C.lxw_chart_line) &C.lxw_chart {
	C.chart_set_high_low_lines(c, line)
	return unsafe { c }
}

// set_series_overlap set the `overlap`(-100 to 100) between series in a Bar/Column chart
pub fn (c &C.lxw_chart) set_series_overlap(overlap i8) &C.lxw_chart {
	C.chart_set_series_overlap(c, overlap)
	return unsafe { c }
}

// set_series_gap set the `gap`(0 to 500) between series in a Bar/Column chart
pub fn (c &C.lxw_chart) set_series_gap(gap u16) &C.lxw_chart {
	C.chart_set_series_gap(c, gap)
	return unsafe { c }
}

// show_blanks_as set the option for displaying blank data in a chart
pub fn (c &C.lxw_chart) show_blanks_as(option Lxw_chart_blank) &C.lxw_chart {
	C.chart_show_blanks_as(c, u8(option))
	return unsafe { c }
}

// show_hidden_data display data on charts from hidden rows or columns
pub fn (c &C.lxw_chart) show_hidden_data() &C.lxw_chart {
	C.chart_show_hidden_data(c)
	return unsafe { c }
}

// set_rotation set the Pie/Doughnut chart `rotation`(The angle of rotation)
pub fn (c &C.lxw_chart) set_rotation(rotation u16) &C.lxw_chart {
	C.chart_set_rotation(c, rotation)
	return unsafe { c }
}

// set_hole_size set the Doughnut chart hole `size`(The hole size as a percentage)
pub fn (c &C.lxw_chart) set_hole_size(size u8) &C.lxw_chart {
	C.chart_set_hole_size(c, size)
	return unsafe { c }
}
