# coding: UTF-8

require 'spec_helper'

describe CellRange do
  let(:book) { TestFile.テーブル }
  let(:sheet1) { book.Sheet1 }
  let(:sheet3) { book.Sheet3 }
  subject { sheet3.A1_C12 }
  let(:table2) { sheet1.A1_C3 }

  after do
    TestFile.close
  end

  describe '#all' do
    it 'でデータの行が取得されています。' do
      subject.all.size.should == 11
      table2.all.size.should == 2
    end
    
    it 'で取得したデータ行からセルのデータが取得されています。' do
      subject.all.first.float.should == 4.4
      subject.all.first.string.should == 'か'
    end

    it 'で空行は取得されません。' do
      sheet1.A1_C5.all.size.should == 2
    end
    
     it 'で空行の次の有効行は取得されます。' do
      sheet1.A1_C6.all.size.should == 3
    end
    it 'で列として指定したデータの行が取得されています。' do
      sheet3.A_C.all.size.should == 11
    end
  end
  
  describe '#where' do
    describe 'で整数列に対して' do
      it '値を指定してデータの検索ができています。' do
        result = subject.where(int: 5)
        result.size.should == 2
        result.should be_all{ |row| row.int == 5 }
      end
     it '範囲を指定してデータの検索ができています。' do
        result = subject.where(int: 2..4)
        result.size.should == 3
        result.should be_all{ |row| 2 <= row.int && row.int <= 4 }
      end
      
      it '値を指定してデータの検索ができています。' do
        result = subject.where(int: 5)
        result.size.should == 2
        result.should be_all{ |row| row.int == 5 }
      end
      it '範囲を指定してデータの検索ができています。' do
        result = subject.where(int: 2..4)
        result.size.should == 3
        result.should be_all{ |row| 2 <= row.int && row.int <= 4 }
      end
      it '複数の要素を指定してデータの検索ができています。' do
        result = subject.where(int: [1, 3, 5])
        result.size.should == 4
        result.should be_all{ |row| [1, 3, 5].include? row.int }
      end
      
      it '複数の条件を指定してデータの検索ができています。' do
        result = subject.where(int: 5, float: 6.6)
        result.size.should == 1
        result.first.int.should == 5
        result.first.float.should == 6.6
      end
    end
    describe 'で実数列に対して' do
      it '値を指定してデータの検索ができています。' do
        result = subject.where(float: 6.6)
        result.size.should == 2
        result.should be_all{ |row| row.float == 6.6 }
      end
      it '範囲を指定してデータの検索ができています。' do
        result = subject.where(float: 2..5)
        result.size.should == 3
        result.should be_all{ |row| 2 <= row.float && row.float <= 5 }
      end
      
      it '値を指定してデータの検索ができています。' do
        result = subject.where(float: 6.6)
        result.size.should == 2
        result.should be_all{ |row| row.float == 6.6 }
      end
      it '範囲を指定してデータの検索ができています。' do
        result = subject.where(float: 2..5)
        result.size.should == 3
        result.should be_all{ |row| 2 <= row.float && row.float <= 5 }
      end
      it '複数の要素を指定してデータの検索ができています。' do
        result = subject.where(float: [1.1, 3.3, 6.6])
        result.size.should == 4
        result.should be_all{ |row| [1.1, 3.3, 6.6].include? row.float }
      end
    end
  end
  
  describe '#order' do
    it 'で並べ替えができます。' do
      subject.order('float').map{|row| row.float }.should == subject.all.map{|row| row.float }.sort
    end
  end
  
  describe '#column_names' do
    it 'で列名が取得されています' do
      subject.column_names.should == [:int, :float, :string]
    end
  end
end