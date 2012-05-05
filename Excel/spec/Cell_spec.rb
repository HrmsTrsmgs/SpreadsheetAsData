# -*- encoding: utf-8 -*-
require File.dirname(__FILE__) + '/../lib/WorkBook'
describe Cell do
	def test_file(file_name)
		return '.\spec\test_data\\' + file_name + '.xlsx'
	end
	
	after(:all) do
		book.close
	end
	
	let(:book) { WorkBook.open(test_file('Book1')) }
	let(:sheet) { book.sheets[2] } 
	
	describe '#value' do
		it 'が数値を取得できる。' do
			sheet.cell(:A1).value.should == 1
			sheet.cell(:B1).value.should == 2
		end
		it 'がtrueを取得できる。' do
			sheet.cell(:A2).value.should == true
		end
		it 'がfalseを取得できる。' do
			sheet.cell(:B2).value.should == false
		end
	end
	
end