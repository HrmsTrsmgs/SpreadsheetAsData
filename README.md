	$: <<  './lib'
	require './lib/work_book'
	
	
	WorkBook.open('./Book1.xlsx') do |book|
	  s = book.Sheet1
	  puts s.A2 + ' ' + s.B2
	  puts s.C2 + s.C3 + s.C4
	  
	  table = s.A1_D7
	  
	  table.where(good_language: 'Ruby').each do |student|
	    puts student.first_name + ' ' + student.last_name
	  end
	  
	  puts table.
	    all.
	    map{|student| student.age }.
	    inject(0){|total, age| total += age } /
	    table.all.size
	end