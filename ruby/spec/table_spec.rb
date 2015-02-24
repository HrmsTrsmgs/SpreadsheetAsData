# coding: UTF-8

require 'spec_helper'
class Table 
end
describe Table do
  let(:book) { TestFile.テーブル }
  let(:sheet) { book.Sheet2 }
  subject { sheet.sheet_tables[0] }
  let(:table2) { sheet.sheet_tables[1] }
  after do
    TestFile.close
  end

  describe '#ref' do
    it 'でセルの名称が取得されています。' do
      #subject.
    end
  end
end