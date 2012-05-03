# -*- encoding: utf-8 -*-
require File.dirname(__FILE__) + '/../lib/WorkBook'
describe WorkSheet do
	def test_file(file_name)
		return '.\spec\test_data\\' + file_name + '.xlsx'
	end
	
	after(:all) do
		book.close
	end
	
	subject { book.Sheet1}
	let(:book) { WorkBook.open(test_file('Book1')) }
	let(:sheet2) { book.Sheet2 } 
	
	describe '#name' do
		it 'でシート名が取得できる。' do
			subject.name.should == 'Sheet1'
			sheet2.name.should == 'Sheet2'
		end
	end
	
	describe '#cell' do
		it 'はセルの値を取得する。' do
			subject.cell('C3').should == 4
		end
		
		it 'は列方向に正しいセルを取得する。' do
			subject.cell('A1').should == 1
			subject.cell('C1').should == 2
		end
		
		it 'は行方向に正しいセルを取得する。' do
			subject.cell('A1').should == 1
			subject.cell('A3').should == 3
		end
		
		it 'はシートを区別して正しいセルを取得する。' do
			subject.cell('A1').should == 1
			sheet2.cell('A1').should == 5
		end
	end
end