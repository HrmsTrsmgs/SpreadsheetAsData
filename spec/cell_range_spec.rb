# coding: UTF-8

require 'spec_helper'

describe CellRange do
  let(:book) { TestFile.book1 }
  let(:sheet) { book.Sheet1 }
  subject { sheet.range('A1', 'C3') }

  after do
    TestFile.close
  end

  describe '#sheet' do
    it 'で所属するシートが取得できます。' do
      subject.sheet.should equal sheet
    end
  end

  describe '#book' do
    it 'で所属するシートが取得できます。' do
      subject.book.should equal book
    end
  end

  describe '#ref' do
    it 'がセル参照の名称を取得できます。' do
      subject.ref.should == 'A1:C3'
      sheet.range('B2', 'B2').ref.should == 'B2:B2'
    end
  end
  
  describe '#to_s' do
    it 'がセルの名称を取得できます。' do
      subject.to_s.should == 'A1:C3'
      sheet.range('B2', 'B2').to_s.should == 'B2:B2'
    end
  end

  describe '#left_upper_cell' do
    it 'が左上のセルを取得できます。' do
      subject.left_upper_cell.should equal sheet.cell(:A1)
      sheet.range('B2', 'B2').left_upper_cell.should equal sheet.cell(:B2)
    end
  end
end