# -*- encoding: Shift_JIS -*- 
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
	
	describe '#sheet' do
		it '�擾�ł���' do
			sheet.cell(:A1).sheet.should equal sheet
		end
	end
	
	describe '#book' do
		it '�擾�ł���' do
			sheet.cell(:A1).book.should equal book
		end
	end
	
	describe '#value' do
		it '�����l���擾�ł���B' do
			sheet.cell(:A1).value.should == 1
			sheet.cell(:B1).value.should == 2
		end
		it '��true���擾�ł���B' do
			sheet.cell(:A2).value.should == true
		end
		it '��false���擾�ł���B' do
			sheet.cell(:B2).value.should == false
		end
		it '����������擾�ł���B' do
			sheet.cell(:A3).value.should == '����������'
			sheet.cell(:B3).value.should == '����������'
		end
	end
	
end