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
      sheet1.cell('C3').ref.should == 'C3'
    end
    it 'はシンボルを渡しても動作します。' do
      sheet1.cell(:C3).ref.should == 'C3'
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
    it 'は一文字の列名称の列を取得します。' do
      sheet1.cell('A1').ref.should == 'A1'
      sheet1.cell('Z1').ref.should == 'Z1'
    end
    it 'は2文字の列名称の列を取得します。' do
      sheet1.cell('AA1').ref.should == 'AA1'
      sheet1.cell('ZZ1').ref.should == 'ZZ1'
    end
    it 'は3文字の列名称の列を取得します。' do
      sheet1.cell('AAA1').ref.should == 'AAA1'
      sheet1.cell('XFD1').ref.should == 'XFD1'
    end
    
    it 'は大きすぎる列名を指定した場合にnilを返します。' do
     sheet1.cell('XFE1').should be_nil
     sheet1.cell('ZZZ1').should be_nil
     sheet1.cell('AAAA1').should be_nil
    end
    
    it 'は存在する大きな行番号のの行を取得します。' do
      sheet1.cell('A1048576').ref.should == 'A1048576'
    end
    
    it 'は大きすぎる行番号を指定した場合にnilを返します。' do
     sheet1.cell('A1048577').should be_nil
    end
  end
  
  describe '#range' do
    describe 'を２引数で呼び出した場合' do
      it 'に範囲を取得できます。' do
        sheet1.range('A1', 'C3').to_s.should == 'A1:C3'
        sheet1.range('B2', 'B2').to_s.should == 'B2:B2' 
      end
      
      it 'に第一引数にシンボルを渡しても動作します。' do
        sheet1.range(:A1, 'C3').to_s.should == 'A1:C3'
      end
      
      it 'に第二引数にシンボルを渡しても動作します。' do
        sheet1.range('A1', :C3).to_s.should == 'A1:C3'
      end
      
      it 'に第一引数と第二引数にシンボルを渡しても動作します。' do
        sheet1.range(:A1, :C3).to_s.should == 'A1:C3'
      end
      
      it 'に同じ範囲を指定した場合に同じセルを返します。' do
        sheet1.range(:A1, :C3).should equal sheet1.range(:A1, :C3)
      end
      
      it 'に第一引数を文字列で呼び出してもシンボルで呼び出しても同じ範囲を指定した場合に同じセルを返します。' do
        sheet1.range('A1', :C3).should equal sheet1.range(:A1, :C3)
      end
      
      it 'に第二引数を文字列で呼び出してもシンボルで呼び出しても同じ範囲を指定した場合に同じセルを返します。' do
        sheet1.range(:A1, 'C3').should equal sheet1.range(:A1, :C3)
      end
      
      it 'は存在する大きな列名の列を取得します。' do
        sheet1.range('XFD1', 'XFD1').ref.should == 'XFD1:XFD1'
      end
      
      it 'は第一引数に大きすぎる列名を指定した場合にnilを返します。' do
       sheet1.range('XFE1', 'A1').should be_nil
      end
      
      it 'は第二引数大きすぎる列名を指定した場合にnilを返します。' do
       sheet1.range('A1', 'XFE1').should be_nil
      end
      
      it 'は存在する大きな行番号のの行を取得します。' do
        sheet1.range('A1048576', 'A1048576').ref.should == 'A1048576:A1048576'
      end
      
      it 'は第一引数に大きすぎる行番号を指定した場合にnilを返します。' do
       sheet1.range('A1048577', 'A1').should be_nil
      end
      
      it 'は第二引数大きすぎる行番号を指定した場合にnilを返します。' do
       sheet1.range('A1', 'A1048577').should be_nil
      end
    end
    describe 'を１引数で呼び出した場合' do
      it 'に範囲を取得できます。' do
        sheet1.range('A1:C3').to_s.should == 'A1:C3'
        sheet1.range('B2:B2').to_s.should == 'B2:B2' 
      end
      it 'にA1_A1形式の文字列を渡して範囲を取得できます。' do
        sheet1.range('A1_C3').to_s.should == 'A1:C3'
        sheet1.range('B2_B2').to_s.should == 'B2:B2' 
      end
      it 'にシンボルを渡しても取得できます。' do
        sheet1.range(:A1_C3).to_s.should == 'A1:C3'
        sheet1.range(:B2_B2).to_s.should == 'B2:B2' 
      end
      it 'に同じ範囲を指定した場合に同じセルを返します。' do
        sheet1.range('A1:C3').should equal sheet1.range('A1:C3')
      end
      it 'に文字列で呼び出してもシンボルで呼び出しても同じ範囲を指定した場合に同じセルを返します。' do
        sheet1.range('A1:C3').should equal sheet1.range(:A1_C3)
      end
      it 'は存在する大きな列名の列を取得します。' do
        sheet1.range('XFD1:XFD1').ref.should == 'XFD1:XFD1'
      end
      
      it 'は一つ目のセル指定に大きすぎる列名を指定した場合にnilを返します。' do
       sheet1.range('XFE1:A1').should be_nil
      end
      
      it 'は二つ目のセル指定大きすぎる列名を指定した場合にnilを返します。' do
       sheet1.range('A1:XFE1').should be_nil
      end
      
      it 'は存在する大きな行番号のの行を取得します。' do
        sheet1.range('A1048576:A1048576').ref.should == 'A1048576:A1048576'
      end
      
      it 'は一つ目のセル指定に大きすぎる行番号を指定した場合にnilを返します。' do
       sheet1.range('A1048577:A1').should be_nil
      end
      
      it 'は二つ目のセル指定に大きすぎる行番号を指定した場合にnilを返します。' do
       sheet1.range('A1:A1048577').should be_nil
      end
    end
    
    it 'の呼び出し引数が多かった場合にArgumentErrorが発生します。' do
      ->{ sheet1.range('', '', '') }.should raise_error ArgumentError, 'wrong number of arguments (3 for 1..2)'
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
  describe '#範囲名称' do
    it 'は範囲を取得できます。' do
      sheet1.A1_C3.should equal sheet1.range('A1:C3')
      sheet1.B2_B2.should equal sheet1.range('B2:B2')
    end
    
    it 'は列の順序が反転していても同じ範囲を取得できます。' do
      sheet1.C1_A3.should equal sheet1.A1_C3
    end
    
    it 'は行の順序が反転していても同じ範囲を取得できます。' do
      sheet1.A3_C1.should equal sheet1.A1_C3
    end
    
    it 'は存在する大きな列名の列を取得します。' do
      sheet1.XFD1_XFD1.ref.should == 'XFD1:XFD1'
    end
    
    it 'は一つ目のセル指定に大きすぎる列名を指定した場合にNoMethodErrorを返します。' do
     ->{ sheet1.XFE1_A1 }.should raise_error NoMethodError
    end
    
    it 'は二つ目のセル指定大きすぎる列名を指定した場合にNoMethodErrorを返します。' do
     ->{ sheet1.A1_XFE1 }.should raise_error NoMethodError
    end
    
    it 'は存在する大きな行番号のの行を取得します。' do
      sheet1.A1048576_A1048576.ref.should == 'A1048576:A1048576'
    end
    
    it 'は一つ目のセル指定に大きすぎる行番号を指定した場合にNoMethodErrorを返します。' do
     ->{ sheet1.A1048577_A1 }.should raise_error NoMethodError
    end
    
    it 'は二つ目のセル指定に大きすぎる行番号を指定した場合にNoMethodErrorを返します。' do
     ->{ sheet1.A1_A1048577 }.should raise_error NoMethodError
    end
  end
end