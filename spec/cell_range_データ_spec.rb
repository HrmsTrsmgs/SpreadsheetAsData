# coding: UTF-8

require 'spec_helper'

describe CellRange do
  let(:book) { TestFile.データ }
  let(:sheet) { book.Sheet1 }
  
  subject { sheet.A1_C3 }
  let(:table2) { sheet.B6_E9 }

  after do
    TestFile.close
  end

  describe '#all' do
    it 'でデータの行が取得されています。' do
      subject.all.size.should == 2
      table2.all.size.should == 3
    end
    
    it 'で取得したデータ行からセルのデータが取得されています。' do
      subject.all.first.float.should == 1.1
      subject.all.first.string.should == 'あいうえお'
    end
  end
  
  describe '#where' do
    it 'でデータの検索ができています。' do
      subject.where(float: 2.2).size.should == 1
      subject.where(float: 2.2).first.float.should == 2.2
    end
  end
  
  describe '#column_names' do
    it 'で列名が取得されています' do
      subject.column_names.should == [:float, :string, :bool]
    end
  end
end