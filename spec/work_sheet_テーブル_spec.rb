# coding:UTF-8

require 'spec_helper'

describe WorkSheet do
  let(:book) { TestFile.テーブル }
  subject { book.Sheet1 }
  let(:sheet2) { book.Sheet2 }

  after do
    TestFile.close
  end

  describe '#defined_name' do
    it 'は参照しているすべてのテーブルを取得できます。' do
      subject.defined_name.size.should == 2
      sheet2.defined_name.size.should == 3
    end
  end
end