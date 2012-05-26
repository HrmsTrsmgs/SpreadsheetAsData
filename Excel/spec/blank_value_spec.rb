# -*- encoding: Shift_JIS -*- 
require 'spec_helper'

describe BlankValue do
  subject { sheet.B1 }
  let(:book) { WorkBook.open test_file('Book1') }
  let(:sheet) { book.Sheet1 } 

  after(:all) do
    book.close
  end

  describe '#==' do
    it '�Ŕ�r�����""�Ɠ����Ƃ������ƂɂȂ�B' do
      subject.should == ''
    end

    it '�Ŕ�r�����0�Ɠ����Ƃ������ƂɂȂ�B' do
      subject.should == 0
    end

    it '�Ŕ�r�����0�ȊO�̐����Ƃ͂Ɠ����Ƃ������Ƃɂ�Ȃ��B' do
      subject.should_not == 1
      subject.should_not == -1
    end

    it '�Ŕ�r�����0�ȊO�̏����Ƃ͂Ɠ����Ƃ������Ƃɂ�Ȃ��B' do
      subject.should_not == 0.01
      subject.should_not == -0.01
    end

    it '�Ŕ�r�����''�ȊO�̕�����Ƃ͂Ɠ����Ƃ������Ƃɂ�Ȃ��B' do
      subject.should_not == 'a'
      subject.should_not == '������'
    end
  end

  describe '#+' do
    it '��0�����Ő����Ƒ������Ƃ��ł���B' do
      (subject + 10).should == 10
      (subject + -5).should == -5
    end

    it '��0�����Ő����ɑ�����邱�Ƃ��ł���B' do
      (10 + subject).should == 10
      (-5 + subject).should == -5
    end

    it '��""�����ŕ�����Ƒ������Ƃ��ł���B' do
      (subject + 'abc').should == 'abc'
    end

    it '��""�����ŕ�����ɑ�����邱�Ƃ��ł���B' do
      ('abc' + subject).should == 'abc'
    end
  end

  describe '#*' do
    it '��0�����Ő����Ɗ|���邱�Ƃ��ł���B' do
      (subject * 10).should == 0
    end

    it '��0�����Ő����Ɋ|�����邱�Ƃ��ł���B' do
      (10 * subject).should == 0
    end

    it '��""�����ŕ�����ɑ�����邱�Ƃ��ł���B' do
      ('abc' * subject).should == ''
    end
  end

  describe '���l���Z' do
    it '�ɗ��p���邱�Ƃ��ł���B' do
      Math::sin(subject).should == 0
      Math::cos(subject).should == 1
    end
  end

  describe '#inspect' do
    it '��{blank}��Ԃ�' do
      subject.inspect.should == '{blank}'
    end
  end

  describe '#to_s' do
    it '��""��Ԃ�' do
      subject.to_s.should == ''
    end
  end
end