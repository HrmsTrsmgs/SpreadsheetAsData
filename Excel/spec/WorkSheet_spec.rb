# -*- encoding: Shift_JIS -*- 
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
		it '�ŃV�[�g�����擾�ł���B' do
			sheet1.name.should == 'Sheet1'
			sheet2.name.should == 'Sheet2'
		end
		
		it '�œ��{��Ŏw�肵���V�[�g����Shift_JIS�Ŏ擾�ł���B' do
			data.name.should == '���낢��ȃf�[�^'
		end
	end
	
	describe '#book' do
		it '�擾�ł���' do
			sheet1.book.should equal book
		end
	end
	
	describe '#cell_value' do
		it '�̓Z���̒l���擾����B' do
			sheet1.cell_value('C3').should == 4
		end
		
		it '�̓V���{����n���Ă����삷��B' do
			sheet1.cell_value(:C3).should == 4
		end
		
		it '�͑��݂��Ȃ��Z�������w�肵������nil��Ԃ�' do
			sheet1.cell_value('a1').should be_nil
		end
	end
	
	describe '#cell' do
		it '�̓Z�����擾����B' do
			sheet1.cell('C3').value.should == 4
		end
		it '�̓V���{����n���Ă����삷��B' do
			sheet1.cell(:C3).value.should == 4
		end
		
		it '�͑��݂��Ȃ��Z�������w�肵������nil��Ԃ�' do
			sheet1.cell('a1').should be_nil
		end
		
		it '�͓����Z���̏ꍇ�͓����I�u�W�F�N�g���擾����B' do
			sheet1.cell('C3').should equal sheet1.cell('C3')
		end                                                                                                                                                           
	end
	
	describe '#�Z����' do
		it '�͒l���擾�ł���B' do
			sheet1.A1.should == 1
			sheet1.C3.should == 4
		end
		it '�͑��݂��Ȃ��Z�������w�肵������NoMethodError��Ԃ�' do
			->{ sheet1.a1 }.should raise_error NoMethodError
		end
		it '�͗�����ɐ������Z�����擾����B' do
			sheet1.A1.should == 1
			sheet1.C1.should == 2
		end
		
		it '�͍s�����ɐ������Z�����擾����B' do
			sheet1.A1.should == 1
			sheet1.A3.should == 3
		end
		
		it '�̓V�[�g����ʂ��Đ������Z�����擾����B' do
			sheet1.A1.should == 1
			sheet2.A1.should == 5
		end
		
		it '�̓V���{����n���Ă����삷��B' do
			sheet1.C3.should == 4
		end
		
		it '�Ő������擾����ꍇ�A�����Z���̏ꍇ�͓����I�u�W�F�N�g���擾����B' do
			sheet1.C3.should equal sheet1.C3
		end
		
		it '�ŕ�������擾����ꍇ�A�����Z���̏ꍇ�͓����I�u�W�F�N�g���擾����B' do
			data.A3.should equal data.A3
		end
	end
end