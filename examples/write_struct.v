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
