# coding: UTF-8

require 'spec_helper'

describe DataRow do
  let(:book) { TestFile.テーブル }
  subject { book.Sheet1.A1_C3.all.first }

  after do
    TestFile.close
  end

  describe '#cell_value' do
    it 'でセルのデータが取得されています。' do
      subject.cell_value('float').should == 1.1
      subject.cell_value('string').should == 'あいうえお'
    end
    
    it 'でシンボル指定でもセルのデータが取得されています。' do
      subject.cell_value(:float).should == 1.1
      subject.cell_value(:string).should == 'あいうえお'
    end
  end
  
  describe '#カラム名' do
    it 'でセルのデータが取得されています。' do
      subject.float.should == 1.1
      subject.string.should == 'あいうえお'
    end
  end
end