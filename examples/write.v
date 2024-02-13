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
