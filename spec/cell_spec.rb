# coding: UTF-8

require 'spec_helper'

describe Cell do
  let(:book) { TestFile.book1 }
  let(:sheet) { book.いろいろなデータ }
  subject { sheet.cell(:A1) }
  
  let(:written_book) { TestFile.book1_copy }
  let(:written_sheet) { written_book.いろいろなデータ }
  let(:written_cell) { written_sheet.cell(:A1) }

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

  describe '#value' do
    it 'が数値を取得できます。' do
      subject.value.should == 1.1
      sheet.cell(:B1).value.should == 2.2
    end

    it 'がtrueを取得できます。' do
      sheet.cell(:A2).value.should == true
    end

    it 'がfalseを取得できます。' do
      sheet.cell(:B2).value.should == false
    end

    it 'が文字列を取得できます。' do
      sheet.cell(:A3).value.should == 'あいうえお'
      sheet.cell(:B3).value.should == 'かきくけこ'
    end
  end

  describe '#value=' do
    it 'が数値を設定できます。' do
      written_cell.value = 3.3
      written_cell.value.should == 3.3
    end

    it 'がtrueを設定できます。' do
      written_cell.value = true
      written_cell.value.should == true
    end

    it 'がfalseを設定できます。' do
      written_cell.value = false
      written_cell.value.should == false
    end
    
    it 'が文字列を設定できます。' do
      written_cell.value = 'abcde'
      written_cell.value.should == 'abcde'
    end
  end

  describe '#ref' do
    it 'がセル参照の名称を取得できます。' do
      subject.ref.should == 'A1'
      sheet.cell(:B1).ref.should == 'B1'
    end
  end
  
  describe '#row_num' do
    it 'がセルの行番号を取得できます。' do
      subject.row_num.should == 1
      sheet.cell(:A3).row_num.should == 3
    end
  end
  
  describe '#to_s' do
    it 'がセルの名称を取得できます。' do
      subject.to_s.should == 'A1'
      sheet.cell(:B1).to_s.should == 'B1'
    end
  end
end