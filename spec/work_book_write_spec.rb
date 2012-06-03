# coding:UTF-8

require 'fileutils'

require 'spec_helper'

describe WorkBook, 'を保存する時' do

  WRITTEN_BOOK = test_file('書き込み')
=begin
  before do
  	book = WorkBook.open(WRITTEN_BOOK) do |book|
      book.Sheet1.cell(:A1).value = ''
    end
  end
=end
  it 'に整数値の書き込みは保存されている。' do
    book = WorkBook.open(WRITTEN_BOOK) do |book|
      book.Sheet1.cell(:A1).value = 999
    end
    
    WorkBook.open(WRITTEN_BOOK) do |book|
      book.Sheet1.A1.should == 999
    end
  end
  
  it 'にtrueの書き込みは保存されている。' do
    book = WorkBook.open(WRITTEN_BOOK) do |book|
      book.Sheet1.cell(:A1).value = true
    end
    
    WorkBook.open(WRITTEN_BOOK) do |book|
      book.Sheet1.A1.should == true
    end
  end
  
  it 'にfalseの書き込みは保存されている。' do
    book = WorkBook.open(WRITTEN_BOOK) do |book|
      book.Sheet1.cell(:A1).value = false
    end
    
    WorkBook.open(WRITTEN_BOOK) do |book|
      book.Sheet1.A1.should == false
    end
  end
  
  it 'に文字列の書き込みは保存されている。' do
    book = WorkBook.open(WRITTEN_BOOK) do |book|
      book.Sheet1.cell(:A1).value = 'abcde'
    end
    
    WorkBook.open(WRITTEN_BOOK) do |book|
      book.Sheet1.A1.should == 'abcde'
    end
  end
end