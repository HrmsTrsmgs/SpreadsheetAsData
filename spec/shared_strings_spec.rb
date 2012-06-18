# coding: UTF-8

require 'spec_helper'

require 'work_book'

describe SharedStrings do
  let(:book) { TestFile.book1_copy }
  subject { book.shared_strings }
  
  after do
    TestFile.close
  end

  describe '#size' do
    it 'は共有文字列の数を返します。' do
      subject.size.should == 2
    end
  end
  
  describe '#[]' do
    it 'は共有文字列を取得できます。' do
      subject[0].should == 'あいうえお'
      subject[1].should == 'かきくけこ'
    end
    
    it 'は追加した共有文字列を取得できます。' do
      i = subject << 'さしすせそ'
      subject[i].should == 'さしすせそ'
    end
  end
  
  describe '#<<' do
    it 'は共有文字列の追加ができます。' do
      subject << 'さしすせそ'
      subject.size.should == 3
    end
    
    it 'は追加された位置が取得できます。' do
      (subject << 'さしすせそ').should == 2
      (subject << 'たちつてと').should == 3
    end
    
    it 'は追加しようとした文字列がすでにあった場合、既存の文字列を利用します。' do
      (subject << 'かきくけこ').should == 1
      (subject << 'あいうえお').should == 0
    end
    
    it 'は既存の文字列を利用したばあい、文字列を追加しません。' do
      subject << 'あいうえお'
      subject.size.should == 2
    end
  end
end