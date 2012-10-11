$: <<  './lib'
require './lib/work_book'


WorkBook.open('./Data.xlsx') do |book|
  s = book.Sheet1
  one = s.A1_A10.all.map{|r| r.one }
  two = s.B1_B10.all.map{|r| r.two }
  three = s.C1_C10.all.map{|r| r.three }
  all = []
  one.each do |on|
    two.each do |tw|
      three.each do |th|
        all << [on, tw, th]
      end
    end
  end
  File.open('./result.csv', 'w') do |f|
    all.shuffle.each do |row|
       f.puts row.join(',')
    end
  end
end