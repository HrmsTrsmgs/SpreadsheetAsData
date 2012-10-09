# coding: UTF-8

require 'spec_helper'

describe CellRange do
  let(:book) { TestFile.テーブル }
  subject { book.Sheet3.A1_C12 }
  let(:table2) { book.Sheet1.A1_C3 }

  after do
    TestFile.close
  end

  describe '#all' do
    it 'でデータの行が取得されています。' do
      subject.all.size.should == 11
      table2.all.size.should == 2
    end
    
    it 'で取得したデータ行からセルのデータが取得されています。' do
      subject.all.first.float.should == 1.1
      subject.all.first.string.should == 'あ'
    end
  end
  
  describe '#where' do
    it 'でデータの検索ができています。' do
      subject.where(int: 2).size.should == 1
      subject.where(int: 2).first.int.should == 2
    end
  end
  
  describe '#column_names' do
    it 'で列名が取得されています' do
      subject.column_names.should == [:int, :float, :string]
    end
  end
end