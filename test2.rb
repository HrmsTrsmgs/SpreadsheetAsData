$: <<  './lib'
require './lib/work_book'


WorkBook.open('./Book2.xlsx') do |book|
  book.Sheet1.B2 = 2
end