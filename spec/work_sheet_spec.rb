# coding:UTF-8

require 'spec_helper'

describe WorkSheet do
  let(:book) { TestFile.book1 }
  let(:sheet1) { book.Sheet1 }
  let(:sheet2) { book.Sheet2 }
  let(:data) { book.いろいろなデータ } 

  after do
    TestFile.close
  end

  describe '#name' do
    it 'がシート名となります。' do
      sheet1.name.should == 'Sheet1'
      sheet2.name.should == 'Sheet2'
    end

    it 'で日本語で指定したシート名が、呼び出した（ファイル名を記述した）コードのエンコードで取得できます。' do
      data.name.should == 'いろいろなデータ'
    end
  end

  describe '#book' do
    it 'が所属するブックを取得します。' do
      sheet1.book.should equal book
    end
  end

  describe '#cell_value' do
    it 'はセルの値を取得します。' do
      sheet1.cell_value('C3').should == 4
    end

    it 'はシンボルを渡しても動作します。' do
      sheet1.cell_value(:C3).should == 4
    end

    it 'は存在しないセル名を指定した時にnilを返します。' do
      sheet1.cell_value('a1').should be_nil
    end
  end

  describe '#cell' do
    it 'はセルを取得します。' do
      sheet1.cell('C3').value.should == 4
    end
    it 'はシンボルを渡しても動作します。' do
      sheet1.cell(:C3).value.should == 4
    end

    it 'は空のセルも取得します。' do
      sheet1.cell('B1').ref.should == 'B1'
    end

    it 'は存在しないセル名を指定した時にnilを返します。' do
      sheet1.cell('a1').should be_nil
    end

    it 'は同じセルの場合は同じオブジェクトを取得します。' do
      sheet1.cell('C3').should equal sheet1.cell('C3')
    end                                                                                                                                                           
    it 'は文字列で呼び出してもシンボルで呼び出しても同じセルの場合は同じオブジェクトを取得します。' do
      sheet1.cell('C3').should equal sheet1.cell(:C3)
    end
    it 'は空のセルの場合も同じセルの場合は同じオブジェクトを取得します。' do
      sheet1.cell('B1').should equal sheet1.cell('B1')
    end                                                                                                                                                           
  end
  
  describe '#range' do
    it 'は範囲を取得できます。' do
      sheet1.range('A1', 'C3').to_s.should == 'A1:C3'
      sheet1.range('B2', 'B2').to_s.should == 'B2:B2' 
    end
    
    it 'は第一引数にシンボルを渡しても動作します。' do
      sheet1.range(:A1, 'C3').to_s.should == 'A1:C3'
    end
    
    it 'は第二引数にシンボルを渡しても動作します。' do
      sheet1.range('A1', :C3).to_s.should == 'A1:C3'
    end
    
    it 'は第一引数と第二引数にシンボルを渡しても動作します。' do
      sheet1.range(:A1, :C3).to_s.should == 'A1:C3'
    end
    
    it 'は同じ範囲を指定した場合に同じセルを返します。' do
      sheet1.range(:A1, :C3).should equal sheet1.range(:A1, :C3)
    end
    
    it 'は第一引数を文字列で呼び出してもシンボルで呼び出しても同じ範囲を指定した場合に同じセルを返します。' do
      sheet1.range('A1', :C3).should equal sheet1.range(:A1, :C3)
    end
    
    it 'は第二引数を文字列で呼び出してもシンボルで呼び出しても同じ範囲を指定した場合に同じセルを返します。' do
      sheet1.range(:A1, 'C3').should equal sheet1.range(:A1, :C3)
    end
  end

  describe '#セル名' do
    it 'は値を取得できます。' do
      sheet1.A1.should == 1
      sheet1.C3.should == 4
    end

    it 'は存在しないセル名を指定した時にNoMethodErrorを返します。' do
      ->{ sheet1.a1 }.should raise_error NoMethodError
    end

    it 'は列を区別して正しいセルを取得します。' do
      sheet1.A1.should == 1
      sheet1.C1.should == 2
    end

    it 'は行を区別して正しいセルを取得します。' do
      sheet1.A1.should == 1
      sheet1.A3.should == 3
    end

    it 'はシートを区別して正しいセルを取得します。' do
      sheet1.A1.should == 1
      sheet2.A1.should == 5
    end

    it 'は空白セルを取得します。' do
      sheet1.B1.should_not be_nil
    end

    it 'で数字を取得する場合、同じセルの場合は同じオブジェクトを取得します。' do
      sheet1.C3.should equal sheet1.C3
    end

    it 'で文字列を取得する場合、同じセルの場合は同じオブジェクトを取得します。' do
      data.A3.should equal data.A3
    end
  end
end