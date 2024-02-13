module vxlsx

@[typedef]
struct C.lxw_format {}

// Lxw_format_underlines underline values for format.set_underline()
pub enum Lxw_format_underlines {
	underline_none              = 0
	underline_single // Single underline
	underline_double // Double underline
	underline_single_accounting // Single accounting underline
	underline_double_accounting // Double accounting underline
}

// Lxw_format_scripts superscript and subscript values for format.set_font_script()
pub enum Lxw_format_scripts {
	font_superscript = 1 // Superscript font
	font_subscript // Subscript font
}

// Lxw_format_alignments alignment values for format.set_align()
pub enum Lxw_format_alignments {
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

// Lxw_format_diagonal_types Diagonal border types
pub enum Lxw_format_diagonal_types {
	diagonal_border_up      = 1 // Cell diagonal border from bottom left to top right
	diagonal_border_down // Cell diagonal border from top left to bottom right
	diagonal_border_up_down // Cell diagonal border in both directions
}

// Lxw_defined_colors Predefined values for common colors
pub enum Lxw_defined_colors {
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

// Lxw_format_patterns pattern value for use with format.set_pattern()
pub enum Lxw_format_patterns {
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

// Lxw_format_borders cell border styles for use with format.set_border()
pub enum Lxw_format_borders {
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

// ===========format functions===========
fn C.format_set_font_name(format &C.lxw_format, font_name &char)
fn C.format_set_font_size(format &C.lxw_format, size f64)
fn C.format_set_font_color(format &C.lxw_format, color u32)
fn C.format_set_bold(format &C.lxw_format)
fn C.format_set_italic(format &C.lxw_format)
fn C.format_set_underline(format &C.lxw_format, style u8)
fn C.format_set_font_strikeout(format &C.lxw_format)
fn C.format_set_font_script(format &C.lxw_format, style u8)
fn C.format_set_num_format(format &C.lxw_format, num_format &char)
fn C.format_set_num_format_index(format &C.lxw_format, index u8)
fn C.format_set_unlocked(format &C.lxw_format)
fn C.format_set_hidden(format &C.lxw_format)
fn C.format_set_align(format &C.lxw_format, alignment u8)
fn C.format_set_text_wrap(format &C.lxw_format)
fn C.format_set_rotation(format &C.lxw_format, angle i16)
fn C.format_set_indent(format &C.lxw_format, level u8)
fn C.format_set_shrink(format &C.lxw_format)
fn C.format_set_pattern(format &C.lxw_format, index u8)
fn C.format_set_bg_color(format &C.lxw_format, color u32)
fn C.format_set_fg_color(format &C.lxw_format, color u32)
fn C.format_set_border(format &C.lxw_format, style u8)
fn C.format_set_bottom(format &C.lxw_format, style u8)
fn C.format_set_top(format &C.lxw_format, style u8)
fn C.format_set_left(format &C.lxw_format, style u8)
fn C.format_set_right(format &C.lxw_format, style u8)
fn C.format_set_border_color(format &C.lxw_format, color u32)
fn C.format_set_bottom_color(format &C.lxw_format, color u32)
fn C.format_set_top_color(format &C.lxw_format, color u32)
fn C.format_set_left_color(format &C.lxw_format, color u32)
fn C.format_set_right_color(format &C.lxw_format, color u32)
fn C.format_set_diag_type(format &C.lxw_format, type_ u8)
fn C.format_set_diag_border(format &C.lxw_format, style u8)
fn C.format_set_diag_color(format &C.lxw_format, color u32)
fn C.format_set_quote_prefix(format &C.lxw_format)
fn C.format_set_font_outline(format &C.lxw_format)
fn C.format_set_font_shadow(format &C.lxw_format)
fn C.format_set_font_family(format &C.lxw_format, value u8)
fn C.format_set_font_charset(format &C.lxw_format, value u8)
fn C.format_set_font_scheme(format &C.lxw_format, font_scheme &char)
fn C.format_set_font_condense(format &C.lxw_format)
fn C.format_set_font_extend(format &C.lxw_format)
fn C.format_set_reading_order(format &C.lxw_format, value u8)
fn C.format_set_theme(format &C.lxw_format, value u8)
fn C.format_set_hyperlink(format &C.lxw_format)
fn C.format_set_color_indexed(format &C.lxw_format, value u8)
fn C.format_set_font_only(format &C.lxw_format)

// set_font_name set the font used in the cell
pub fn (f &C.lxw_format) set_font_name(font_name string) &C.lxw_format {
	C.format_set_font_name(f, font_name.str)
	return unsafe { f }
}

// set_font_size set the size of the font used in the cell
pub fn (f &C.lxw_format) set_font_size(size f64) &C.lxw_format {
	C.format_set_font_size(f, size)
	return unsafe { f }
}

// set_font_color set the color of the font used in the cell
pub fn (f &C.lxw_format) set_font_color(color u32) &C.lxw_format {
	C.format_set_font_color(f, color)
	return unsafe { f }
}

// set_bold turn on bold for the format font
pub fn (f &C.lxw_format) set_bold() &C.lxw_format {
	C.format_set_bold(f)
	return unsafe { f }
}

// set_italic turn on italic for the format font
pub fn (f &C.lxw_format) set_italic() &C.lxw_format {
	C.format_set_italic(f)
	return unsafe { f }
}

// set_underline turn on underline for the format
pub fn (f &C.lxw_format) set_underline(style Lxw_format_underlines) &C.lxw_format {
	C.format_set_underline(f, u8(style))
	return unsafe { f }
}

// set_font_strikeout set the strikeout property of the font
pub fn (f &C.lxw_format) set_font_strikeout() &C.lxw_format {
	C.format_set_font_strikeout(f)
	return unsafe { f }
}

// set_font_script set the superscript/subscript property of the font
pub fn (f &C.lxw_format) set_font_script(style Lxw_format_scripts) &C.lxw_format {
	C.format_set_font_script(f, u8(style))
	return unsafe { f }
}

// set_num_format set the number format for a cell
pub fn (f &C.lxw_format) set_num_format(num_format string) &C.lxw_format {
	C.format_set_num_format(f, num_format.str)
	return unsafe { f }
}

// set_num_format_index set the Excel built-in number format for a cell
pub fn (f &C.lxw_format) set_num_format_index(index u8) &C.lxw_format {
	C.format_set_num_format_index(f, index)
	return unsafe { f }
}

// set_unlocked set the cell unlocked state
pub fn (f &C.lxw_format) set_unlocked() &C.lxw_format {
	C.format_set_unlocked(f)
	return unsafe { f }
}

// set_hidden hide formulas in a cell
pub fn (f &C.lxw_format) set_hidden() &C.lxw_format {
	C.format_set_hidden(f)
	return unsafe { f }
}

// set_align set the alignment for data in the cell
pub fn (f &C.lxw_format) set_align(alignment Lxw_format_alignments) &C.lxw_format {
	C.format_set_align(f, u8(alignment))
	return unsafe { f }
}

// set_text_wrap wrap text in a cell
pub fn (f &C.lxw_format) set_text_wrap() &C.lxw_format {
	C.format_set_text_wrap(f)
	return unsafe { f }
}

// set_rotation set the rotation of the text in a cell
pub fn (f &C.lxw_format) set_rotation(angle i16) &C.lxw_format {
	C.format_set_rotation(f, angle)
	return unsafe { f }
}

// set_indent set the cell text indentation level
pub fn (f &C.lxw_format) set_indent(level u8) &C.lxw_format {
	C.format_set_indent(f, level)
	return unsafe { f }
}

// set_shrink turn on the text "shrink to fit" for a cell
pub fn (f &C.lxw_format) set_shrink() &C.lxw_format {
	C.format_set_shrink(f)
	return unsafe { f }
}

// set_pattern set the background fill pattern for a cell
pub fn (f &C.lxw_format) set_pattern(index Lxw_format_patterns) &C.lxw_format {
	C.format_set_pattern(f, u8(index))
	return unsafe { f }
}

// set_bg_color set the pattern background color for a cell
pub fn (f &C.lxw_format) set_bg_color(color u32) &C.lxw_format {
	C.format_set_bg_color(f, color)
	return unsafe { f }
}

// set_fg_color set the pattern foreground color for a cell
pub fn (f &C.lxw_format) set_fg_color(color u32) &C.lxw_format {
	C.format_set_fg_color(f, color)
	return unsafe { f }
}

// set_border set the cell border style
pub fn (f &C.lxw_format) set_border(style Lxw_format_borders) &C.lxw_format {
	C.format_set_border(f, u8(style))
	return unsafe { f }
}

// set_bottom set the cell bottom border style
pub fn (f &C.lxw_format) set_bottom(style Lxw_format_borders) &C.lxw_format {
	C.format_set_bottom(f, u8(style))
	return unsafe { f }
}

// set_top set the cell top border style
pub fn (f &C.lxw_format) set_top(style Lxw_format_borders) &C.lxw_format {
	C.format_set_top(f, u8(style))
	return unsafe { f }
}

// set_left set the cell left border style
pub fn (f &C.lxw_format) set_left(style Lxw_format_borders) &C.lxw_format {
	C.format_set_left(f, u8(style))
	return unsafe { f }
}

// set_right set the cell right border style
pub fn (f &C.lxw_format) set_right(style Lxw_format_borders) &C.lxw_format {
	C.format_set_right(f, u8(style))
	return unsafe { f }
}

// set_border_color set the color of the cell border
pub fn (f &C.lxw_format) set_border_color(color u32) &C.lxw_format {
	C.format_set_border_color(f, color)
	return unsafe { f }
}

// set_bottom_color set the color of the bottom cell border
pub fn (f &C.lxw_format) set_bottom_color(color u32) &C.lxw_format {
	C.format_set_bottom_color(f, color)
	return unsafe { f }
}

// set_top_color set the color of the top cell border
pub fn (f &C.lxw_format) set_top_color(color u32) &C.lxw_format {
	C.format_set_top_color(f, color)
	return unsafe { f }
}

// set_left_color set the color of the left cell border
pub fn (f &C.lxw_format) set_left_color(color u32) &C.lxw_format {
	C.format_set_left_color(f, color)
	return unsafe { f }
}

// set_right_color set the color of the right cell border
pub fn (f &C.lxw_format) set_right_color(color u32) &C.lxw_format {
	C.format_set_right_color(f, color)
	return unsafe { f }
}

// set_diag_type set the diagonal cell border type
pub fn (f &C.lxw_format) set_diag_type(diagonal_type Lxw_format_diagonal_types) &C.lxw_format {
	C.format_set_diag_type(f, u8(diagonal_type))
	return unsafe { f }
}

// set_diag_border set the diagonal cell border style
pub fn (f &C.lxw_format) set_diag_border(style Lxw_format_borders) &C.lxw_format {
	C.format_set_diag_border(f, u8(style))
	return unsafe { f }
}

// set_diag_color set the diagonal cell border color
pub fn (f &C.lxw_format) set_diag_color(color u32) &C.lxw_format {
	C.format_set_diag_color(f, color)
	return unsafe { f }
}

// set_quote_prefix turn on quote prefix for the format
pub fn (f &C.lxw_format) set_quote_prefix() &C.lxw_format {
	C.format_set_quote_prefix(f)
	return unsafe { f }
}

// set_font_outline set the font_outline property
pub fn (f &C.lxw_format) set_font_outline() &C.lxw_format {
	C.format_set_font_outline(f)
	return unsafe { f }
}

// set_font_shadow set the font_shadow property
pub fn (f &C.lxw_format) set_font_shadow() &C.lxw_format {
	C.format_set_font_shadow(f)
	return unsafe { f }
}

// set_font_family set the font_family property
pub fn (f &C.lxw_format) set_font_family(value u8) &C.lxw_format {
	C.format_set_font_family(f, value)
	return unsafe { f }
}

// set_font_charset set the font_charset property
pub fn (f &C.lxw_format) set_font_charset(value u8) &C.lxw_format {
	C.format_set_font_charset(f, value)
	return unsafe { f }
}

// set_font_scheme set the font_scheme property
pub fn (f &C.lxw_format) set_font_scheme(font_scheme string) &C.lxw_format {
	C.format_set_font_scheme(f, font_scheme.str)
	return unsafe { f }
}

// set_font_condense set the font_condense property
pub fn (f &C.lxw_format) set_font_condense() &C.lxw_format {
	C.format_set_font_condense(f)
	return unsafe { f }
}

// set_font_extend set the font_extend property
pub fn (f &C.lxw_format) set_font_extend() &C.lxw_format {
	C.format_set_font_extend(f)
	return unsafe { f }
}

// set_reading_order set the reading_order property
pub fn (f &C.lxw_format) set_reading_order(value u8) &C.lxw_format {
	C.format_set_reading_order(f, value)
	return unsafe { f }
}

// set_theme set the theme property
pub fn (f &C.lxw_format) set_theme(value u8) &C.lxw_format {
	C.format_set_theme(f, value)
	return unsafe { f }
}

// set_hyperlink set the theme property
pub fn (f &C.lxw_format) set_hyperlink() &C.lxw_format {
	C.format_set_hyperlink(f)
	return unsafe { f }
}

// set_color_indexed set the color_indexed property
pub fn (f &C.lxw_format) set_color_indexed(value u8) &C.lxw_format {
	C.format_set_color_indexed(f, value)
	return unsafe { f }
}

// set_font_only set the font_only property
pub fn (f &C.lxw_format) set_font_only() &C.lxw_format {
	C.format_set_font_only(f)
	return unsafe { f }
}
