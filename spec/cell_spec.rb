# coding: UTF-8

require 'spec_helper'

describe Cell do
  let(:book) { WorkBook.open(test_file('Book1')) }
  let(:sheet) { book.いろいろなデータ }

  after(:all) do
    book.close
  end

  describe '#sheet' do
    it 'で所属するシートが取得できる' do
      sheet.cell(:A1).sheet.should equal sheet
    end
  end

  describe '#book' do
    it 'で所属するシートが取得できる' do
      sheet.cell(:A1).book.should equal book
    end
  end

  describe '#value' do
    it 'が数値を取得できる。' do
      sheet.cell(:A1).value.should == 1.1
      sheet.cell(:B1).value.should == 2.2
    end

    it 'がtrueを取得できる。' do
      sheet.cell(:A2).value.should == true
    end

    it 'がfalseを取得できる。' do
      sheet.cell(:B2).value.should == false
    end

    it 'が文字列を取得できる。' do
      sheet.cell(:A3).value.should == 'あいうえお'
      sheet.cell(:B3).value.should == 'かきくけこ'
    end
  end

  describe '#value=' do
    it 'が数値を設定できる。' do
      sheet.cell(:A1).value = 3.3
      sheet.cell(:A1).value.should == 3.3
    end

    it 'がtrueを設定できる。' do
      sheet.cell(:A1).value = true
      sheet.cell(:A1).value.should == true
    end

    it 'がfalseを設定できる。' do
      sheet.cell(:A1).value = false
      sheet.cell(:A1).value.should == false
    end
  end

  describe '#ref' do
    it 'がセル参照の名称を取得できる。' do
      sheet.cell(:A1).ref.should == 'A1'
      sheet.cell(:B1).ref.should == 'B1'
    end
  end
end