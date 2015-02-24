$: <<  './lib'
require './lib/work_book'

WorkBook.open('./Book1.xlsx') do |book|
  sheet = book.Sheet1
  
  p sheet.A2 # "Jacob"
  p sheet.C3 # 10.0
  
  p sheet.E5 # {blank}
  
  p sheet.A2 + sheet.E5 # "Jacob"
  p sheet.C3 * sheet.E5 # 0.0
  
  table = sheet.A1_D7
  
  table.where(age: 12..19).each do |row|
    p row.first_name # "Isabella" "William" "Emma"
  end
end