# -*- encoding: Shift_JIS -*- 

require 'fileutils'

require 'spec_helper'

describe WorkBook do

  BOOK_OPEND_EACH_TIME = '�����J�����Book'

  subject{WorkBook.open(test_file('Book1'))}
  let(:book2){WorkBook.open(test_file('Book2'))}
  let(:anomaly){WorkBook.open(test_file('�ϑ������[�V����'))}
  let(:utf_8){WorkBook.open(test_file('UTF-8�ŊJ��Book'), 'UTF-8')}
  let(:euc_jp){WorkBook.open(test_file('EUC-JP�ŊJ��Book'), 'EUC-JP')}

  after(:all) do
    subject.close
    book2.close
    anomaly.close
    utf_8.close
    euc_jp.close
  end

  describe '#open' do
    it '�̖߂�l��WorkBook�ł���B' do
      subject.class.should == WorkBook
    end

    it '�̃u���b�N������WorkBook�ł���B' do
      WorkBook.open(test_file(BOOK_OPEND_EACH_TIME)) do |book|
        book.class.should == WorkBook
      end
    end

    context '�̓t�@�C�����̃����[�V����' do
      it '���ϑ��I�ȏꍇ�ł����삷��B' do
        anomaly.Sheet1.C3.should == 4
      end
    end

    it '�ŃG���R�[�h���w�肷�邱�Ƃɂ��A�V�[�g�����擾���镶���R�[�h��I���ł���B' do
      utf_8.sheets[2].name.should == '���낢��ȃf�[�^'.encode('UTF-8')
      euc_jp.sheets[2].name.should == '���낢��ȃf�[�^'.encode('EUC-JP')
    end

    it '�ŃG���R�[�h���w�肵�����ɁA�V�[�g��I�����镶����̃G���R�[�h�͐�������Ȃ��B' do
      utf_8['���낢��ȃf�[�^'].should equal utf_8.sheets[2]
      euc_jp['���낢��ȃf�[�^'].should equal euc_jp.sheets[2]
    end

    it '�ŃG���R�[�h���w�肵�����ɁA���\�b�h�Ăяo���̃G���R�[�h�͐�������Ȃ��B' do
      utf_8.���낢��ȃf�[�^.should equal utf_8.sheets[2]
      euc_jp.���낢��ȃf�[�^.should equal euc_jp.sheets[2]
    end

    it '�ŃG���R�[�h���w�肷�邱�ƂŁA������f�[�^���擾���镶���R�[�h���w��ł���B' do
      utf_8.���낢��ȃf�[�^.A3.should == '����������'.encode('UTF-8')
      euc_jp.���낢��ȃf�[�^.A3.should == '����������'.encode('EUC-JP')
    end

    context '�ł̓����t�@�C������' do
      it '�̎��ɉ𓀂��s���Ă���B' do
        temp_dir_name = test_file('tmp_' + BOOK_OPEND_EACH_TIME)
        FileUtils.remove_entry(temp_dir_name) if Dir.exist?(temp_dir_name)
        WorkBook.open(test_file(BOOK_OPEND_EACH_TIME)) do |book|
          Dir.exist?(temp_dir_name).should == true
        end
      end

      it '�̏I�����ɉ𓀂�����ƃt�@�C���̍폜���s���Ă���B' do
        temp_dir_name = test_file('tmp_' + BOOK_OPEND_EACH_TIME)
        FileUtils.remove_entry(temp_dir_name) if Dir.exist?(temp_dir_name)
        WorkBook.open(test_file(BOOK_OPEND_EACH_TIME)) {|book| }
        Dir.exist?(temp_dir_name).should == false
      end
    end
  end

  describe '#file_path' do
    it '�Ŏw�肵���p�X�擾�ł���B' do
      subject.file_path.should == test_file('Book1')
      book2.file_path.should == test_file('Book2')
    end
    it '���Aopen���ɃX���b�V����؂�Ńp�X���w�肵���ꍇ�ɂ��A�w�肵���ʂ�Ƀp�X�擾�ł���B' do
      test_file_slash = test_file(BOOK_OPEND_EACH_TIME).gsub('\\', '/')
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
    context '�̈����Ƃ��ĕ�������w�肵�����B' do
      it '�ɃV�[�g���擾�ł���' do
        subject['Sheet1'].should equal subject.sheets[0]
        subject['Sheet2'].should equal subject.sheets[1]
      end

      it '�ɑ��݂��Ȃ����O���w�肳�ꂽ���ɂ�nil��Ԃ��B' do
        subject['Sheet999'].should be_nil
      end

      it '�ɓ���V�[�g�𓯈�I�u�W�F�N�g�Ƃ��Ĉ����B' do
        subject['Sheet1'].should equal subject['Sheet1']
      end

      it '�ɓ��{�ꖼ�̃V�[�g���擾�ł���B' do
        subject['���낢��ȃf�[�^'].should equal subject.sheets[2]
      end
    end
    context '�̈����Ƃ��ăV���{�����w�肵����' do
      it '�ɃV�[�g���擾�ł���B' do
        subject[:Sheet1].should equal subject.sheets[0]
        subject[:Sheet2].should equal subject.sheets[1]
      end
    end
  end

  describe '#�V�[�g��' do
    it '���V�[�g���擾�ł���B' do
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