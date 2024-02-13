# vxlsx
## Description
A vlang libxlsxwriter module

This is an effort port libxlsxwriter(https://github.com/jmcnamara/libxlsxwriter) to vlang.

## Simple Usage

vxlsx provide a simple function interface.

```v
module main

import vxlsx

struct Expense {
	item string
	cost int
}

fn main() {
	w := vxlsx.new_workbook('tutorial01.xlsx')
	s := w.add_worksheet('test')

	expenses := [Expense{'Rent', 1000}, Expense{'Gas', 100}, Expense{'Food', 300},
		Expense{'Gym', 50}]

	mut col := u16(0)

	for row in 0 .. 4 {
		s.write(row, col, expenses[row].item) // `write` can write multiple type of data
		s.write(row, col + 1, expenses[row].cost)
	}
	s.write(4, col, 'total')
	s.write(4, col + 1, '=SUM(B1:B4)') // `write` can write formula, which starts with '='
	w.close()
}

```

The `write` function can write multiple type of data.
`write` can write formula too, which starts with `=`.

## Write a struct

vxlsx provide a simple method write struct.
```v
module main

import vxlsx

struct Expense {
	item string              @[row: 0; title: 'MyItems'] // use attr as directive
	cost vxlsx.All_Data_Type @[title: 'MyCost']
}

fn main() {
	w := vxlsx.new_workbook('tutorial02.xlsx')
	s := w.add_worksheet('test')

	expenses := [Expense{'Rent', 1000}, Expense{'Gas', 100}, Expense{'Food', 300},
		Expense{'Gym', 50}, Expense{'total', '=SUM(B2:B5)'}]
	s.write_struct(expenses) // `write_struc` use struct attr write the records
	w.close()
}
```

Struct `Expense` first field has attr, which direct how to write the field(s) to xlsx.
Currently, only support 3 attrs:
- `row` set the start row, and every record per col;
- `col` set the start col, and every record per row; (Note: You can not use `row` and `col` both)
- `title` write the title line to xlsx, and set the title for the field;

## TODO

Many enum have not been ported.
