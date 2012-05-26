# -*- encoding: Shift_JIS -*- 

require 'spec_helper'

describe Cell do
  let(:book) { WorkBook.open(test_file('Book1')) }
  let(:sheet) { book.���낢��ȃf�[�^ }

  after(:all) do
    book.close
  end

  describe '#sheet' do
    it '�ŏ�������V�[�g���擾�ł���' do
      sheet.cell(:A1).sheet.should equal sheet
    end
  end

  describe '#book' do
    it '�ŏ�������V�[�g���擾�ł���' do
      sheet.cell(:A1).book.should equal book
    end
  end

  describe '#value' do
    it '�����l���擾�ł���B' do
      sheet.cell(:A1).value.should == 1.1
      sheet.cell(:B1).value.should == 2.2
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

  describe '#ref' do
    it '���Z���Q�Ƃ̖��̂��擾�ł���B' do
      sheet.cell(:A1).ref.should == 'A1'
      sheet.cell(:B1).ref.should == 'B1'
    end
  end
end