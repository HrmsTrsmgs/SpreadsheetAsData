# -*- encoding: utf-8 -*-
require File.dirname(__FILE__) + '/../lib/WorkBook'
describe WorkSheet do
	def test_file(file_name)
		return '.\spec\test_data\\' + file_name + '.xlsx'
	end
	
	after(:all) do
		book.close
	end
	
	let(:sheet1) { book.Sheet1}
	let(:book) { WorkBook.open(test_file('Book1')) }
	let(:sheet2) { book.Sheet2 }
	let(:data) { book.sheets[2] } 
	
	describe '#name' do
		it 'でシート名が取得できる。' do
			sheet1.name.should == 'Sheet1'
			sheet2.name.should == 'Sheet2'
		end
		
		it 'で日本語で指定したシート名がShift_JISで取得できる。' do
			data.name.should == 'いろいろなデータ'.encode("Shift_JIS")
		end
	end
	
	describe '#cell_value' do
		it 'はセルの値を取得する。' do
			sheet1.cell_value('C3').should == 4
		end
		
		it 'は列方向に正しいセルを取得する。' do
			sheet1.cell_value('A1').should == 1
			sheet1.cell_value('C1').should == 2
		end
		
		it 'は行方向に正しいセルを取得する。' do
			sheet1.cell_value('A1').should == 1
			sheet1.cell_value('A3').should == 3
		end
		
		it 'はシートを区別して正しいセルを取得する。' do
			sheet1.cell_value('A1').should == 1
			sheet2.cell_value('A1').should == 5
		end
		
		it 'はシンボルを渡しても動作する。' do
			sheet1.cell_value(:C3).should == 4
		end
	end
	
	describe '#セル名' do
		it 'は存在しないセル名を指定した時にNoMethodErrorを返す' do
			lambda { sheet1.a1 }.should raise_error NoMethodError
		end
		
		context 'が値を取得する時' do
			it 'に数値を取得できる。' do
				sheet1.A1.should == 1
				sheet1.C3.should == 4
			end
			it 'にtrueを取得できる。' do
				data.A1.should == true
			end
			it 'にfalseを取得できる。' do
				data.B1.should == false
			end
		end
	end
end