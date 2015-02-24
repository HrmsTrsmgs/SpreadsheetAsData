# coding:UTF-8

require 'spec_helper'

describe WorkSheet do
  let(:book) { TestFile.テーブル }
  subject { book.Sheet1 }
  let(:sheet2) { book.Sheet2 }

  after do
    TestFile.close
  end
 
  describe '#sheet_tables' do
    it 'は参照しているすべてのテーブルを取得できます。' do
      subject.sheet_tables.size.should == 3
      sheet2.sheet_tables.size.should == 4
    end
  end

  describe '#defined_name' do
    it 'は参照しているすべての定義済み名前を取得できます。' do
      subject.defined_name.size.should == 3
      sheet2.defined_name.size.should == 4
    end
  end
end