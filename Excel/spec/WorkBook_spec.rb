# -*- encoding: Shift_JIS -*- 
require File.dirname(__FILE__) + '/../lib/WorkBook'
describe WorkBook do
	def test_file(file_name)
		return '.\spec\test_data\\' + file_name + '.xlsx'
	end
	
	def unopened_file_name
		return '�����J�����Book'
	end
	
	after(:all) do
		subject.close
		book2.close
		anomaly.close
		utf_8.close
		euc_jp.close
	end
	
	subject{WorkBook.open(test_file('Book1'))}
	let(:book2){WorkBook.open(test_file('Book2'))}
	let(:anomaly){WorkBook.open(test_file('�ϑ������[�V����'))}
	let(:utf_8){WorkBook.open(test_file('UTF-8�ŊJ��Book'), 'UTF-8')}
	let(:euc_jp){WorkBook.open(test_file('EUC-JP�ŊJ��Book'), 'EUC-JP')}
	
	describe '#open' do
	
		it '�̖߂�l��nil�ł͂Ȃ��B' do
			should_not be_nil
		end
		
		it '�̖߂�l��WorkBook�ł���B' do
			subject.class.should == WorkBook
		end
	
		it '�̃u���b�N������nil�ł͂Ȃ��B' do
			WorkBook.open(test_file(unopened_file_name)) do |book|
				book.should_not be_nil
			end
		end
		
		it '�̃u���b�N������WorkBook�ł���B' do
			WorkBook.open(test_file(unopened_file_name)) do |book|
				book.class.should == WorkBook
			end
		end
		
		context '�t�@�C�����̃����[�V����' do
			it '���ϑ��I�ȏꍇ�ł����삷��B' do
				anomaly.Sheet1.C3.should == 4
			end
		end
		context '�����t�@�C������' do
			it '�̎��ɉ𓀂��s���Ă���B' do
				temp_dir_name = test_file('tmp_' + unopened_file_name)
				delete_all(temp_dir_name) if Dir.exist?(temp_dir_name)
				WorkBook.open(test_file(unopened_file_name)) do |book|
					Dir.exist?(temp_dir_name).should == true
				end
			end
			
			it '�̏I�����ɉ𓀂�����ƃt�@�C���̍폜���s���Ă���B' do
				temp_dir_name = test_file('tmp_' + unopened_file_name)
				delete_all(temp_dir_name) if Dir.exist?(temp_dir_name)
				WorkBook.open(test_file(unopened_file_name)) do |book|
				end
				Dir.exist?(temp_dir_name).should == false
			end
			
			it '�ŃG���R�[�h���w�肵�A�V�[�g�����擾���镶���R�[�h��I���ł���B' do
				utf_8.sheets[2].name.should == '���낢��ȃf�[�^'.encode('UTF-8')
				euc_jp.sheets[2].name.should == '���낢��ȃf�[�^'.encode('EUC-JP')
			end
			
			it '�ŃG���R�[�h���w�肵�����ɁA�V�[�g��I�����镶����͐�������Ȃ��B' do
				utf_8['���낢��ȃf�[�^'].should equal utf_8.sheets[2]
				euc_jp['���낢��ȃf�[�^'].should equal euc_jp.sheets[2]
			end
			
			it '�ŃG���R�[�h���w�肵�����ɁA�V�[�g���擾���郁�\�b�h�͐�������Ȃ��B' do
				utf_8.���낢��ȃf�[�^.should equal utf_8.sheets[2]
				euc_jp.���낢��ȃf�[�^.should equal euc_jp.sheets[2]
			end
			
			it '�ŃG���R�[�h���w�肷�邱�ƂŁA������f�[�^���擾���镶���R�[�h���w��ł���B' do
				utf_8.���낢��ȃf�[�^.A3.should == '����������'.encode('UTF-8')
				euc_jp.���낢��ȃf�[�^.A3.should == '����������'.encode('EUC-JP')
			end
		end
	end
	describe '#file_path' do
		it '���w�肵���p�X�擾�ł���B' do
			subject.file_path.should == test_file('Book1')
			book2.file_path.should == test_file('Book2')
		end
		it '���X���b�V����؂�Ŏw�肵���ꍇ�ɂ��w�肵���ʂ�Ƀp�X�擾�ł���B' do
			test_file_slash = test_file(unopened_file_name).gsub('\\', '/')
			WorkBook.open(test_file_slash) do |book|
				book.file_path.should == test_file_slash
			end
		end
	end
	describe '#sheets' do
		it '�ŃV�[�g���擾�ł���B' do
			book2.sheets.map(&:name).should == %w[Sheet1 Sheet2]
		end
	end
	
	describe '#[]' do
		context '�����Ƃ��ĕ�������w��' do
			it '�ŃV�[�g���擾�ł���' do
				subject['Sheet1'].should equal subject.sheets[0]
				subject['Sheet2'].should equal subject.sheets[1]
			end
			
			it '�͑��݂��Ȃ����O���w�肳�ꂽ���ɂ�nil��Ԃ��B' do
				subject['Sheet999'].should be_nil
			end
			
			it '�͓���V�[�g�𓯈�I�u�W�F�N�g�Ƃ��Ĉ����B' do
				subject['Sheet1'].should equal subject['Sheet1']
			end
			
			it '�œ��{�ꖼ�̃V�[�g���擾�ł���' do
				subject['���낢��ȃf�[�^'].should equal subject.sheets[2]
			end
		end
		context '�����Ƃ��ăV���{�����w��' do
			it '�ŃV�[�g���擾�ł���' do
				subject[:Sheet1].should equal subject.sheets[0]
				subject[:Sheet2].should equal subject.sheets[1]
			end
		end
	end
	
	
	describe '#�V�[�g��' do
		it '�ŃV�[�g���擾�ł���' do
			subject.Sheet1.should equal subject.sheets[0]
			subject.Sheet2.should equal subject.sheets[1]
		end
		
		it '�͑��݂��Ȃ��V�[�g�����w�肳�ꂽ���ɂ�NoMethodError��Ԃ��B' do
			->{subject.Sheet999}.should raise_error NoMethodError
		end
		
		it '�œ��{�ꖼ�̃V�[�g���擾�ł���' do
			subject.���낢��ȃf�[�^.should equal subject.sheets[2]
		end
	end
end

# �w�肵���p�X�̃t�H���_�y�сA���̉��ɂ���t�@�C���A�t�H���_�����ׂč폜���܂��B
def delete_all(deleted)
	if FileTest.directory?(deleted) then
		Dir.foreach(deleted) do |child_path|
			next if child_path =~ /^\.+$/
			delete_all(deleted.sub(/\/+$/,"") + "/" + child_path )
		end
		Dir.rmdir(deleted) rescue ""
	else
		File.delete(deleted)
	end
end	