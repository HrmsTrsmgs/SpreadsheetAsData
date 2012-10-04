# coding: UTF-8

require 'spec_helper'

describe CellRange do
  let(:book) { TestFile.データ }
  let(:sheet) { book.Sheet1 }
  
  let(:table1) { sheet.A1_C3 }
  let(:table2) { sheet.B6_E9 }

  after do
    TestFile.close
  end

  describe '#all' do
    it 'でデータの行が取得されています。' do
      table1.all.size.should == 2
      table2.all.size.should == 3
    end
  end
  describe '#column_names' do
    it 'で列名が取得されています' do
      table1.column_names.should == ['float', 'string', 'bool']
    end
  end
end