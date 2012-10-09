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
    describe 'で整数列に対して' do
      it '値を指定してデータの検索ができています。' do
        subject.where(int: 5).size.should == 2
        subject.where(int: 5).should be_all{ |row| row.int == 5 }
      end
     it '範囲を指定してデータの検索ができています。' do
        subject.where(int: 2..4).size.should == 3
        subject.where(int: 2..4).should be_all{ |row| 2 <= row.int && row.int <= 4 }
      end
      
      it '値を指定してデータの検索ができています。' do
        subject.where(int: 5).size.should == 2
        subject.where(int: 5).should be_all{ |row| row.int == 5 }
      end
     it '範囲を指定してデータの検索ができています。' do
        subject.where(int: 2..4).size.should == 3
        subject.where(int: 2..4).should be_all{ |row| 2 <= row.int && row.int <= 4 }
      end
    end
    describe 'で実数列に対して' do
      it '値を指定してデータの検索ができています。' do
        subject.where(float: 6.6).size.should == 2
        subject.where(float: 6.6).should be_all{ |row| row.float == 6.6 }
      end
     it '範囲を指定してデータの検索ができています。' do
        subject.where(float: 2..5).size.should == 3
        subject.where(float: 2..5).should be_all{ |row| 2 <= row.float && row.float <= 5 }
      end
      
      it '値を指定してデータの検索ができています。' do
        subject.where(float: 6.6).size.should == 2
        subject.where(float: 6.6).should be_all{ |row| row.float == 6.6 }
      end
     it '範囲を指定してデータの検索ができています。' do
        subject.where(float: 2..5).size.should == 3
        subject.where(float: 2..5).should be_all{ |row| 2 <= row.float && row.float <= 5 }
      end
    end
  end
  
  describe '#column_names' do
    it 'で列名が取得されています' do
      subject.column_names.should == [:int, :float, :string]
    end
  end
end