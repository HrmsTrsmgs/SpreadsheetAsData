# coding: UTF-8

require 'spec_helper'

describe CellName do
  describe '::valid?' do
    it 'は有効なセルではtrueとなる' do
      CellName.valid?('A1').should be_truthy
    end
    it 'は無効なセルではfalseとなる' do
      CellName.valid?('A').should be_falsey
      CellName.valid?('1').should be_falsey
    end
      it 'はメソッド呼び出しではなく、when句に利用しても使えます。' do
      CellName.valid?.should === 'A1'
      CellName.valid?.should_not === 'A'
    end
  end
  describe 'をハッシュのキーとして利用する時に、CellName' do
    let(:hash) { { CellName.new('A1') => 1 } }
    it 'は位置が同じ時に同値と判定されます。' do      
      hash[CellName.new('A1')].should == 1
    end
    it 'は列が違うときにに同値ではないと判定されます。' do      
      hash[CellName.new('B1')].should be_nil
    end
    it 'は行が違うときにに同値ではないと判定されます。' do      
      hash[CellName.new('A2')].should be_nil
    end
  end
  describe '#initialize' do
    it 'は文字列を受け取ります。' do
      CellName.new('A1').to_s.should == 'A1'
      CellName.new('B2').to_s.should == 'B2'
    end
    it 'はシンボルを受け取ります。' do
      CellName.new(:A1).to_s.should == 'A1'
      CellName.new(:B2).to_s.should == 'B2'
    end
  end
  describe '#column_name' do
    it 'で列名を取得できます' do
      CellName.new('A1').column_name.should == 'A'
      CellName.new('B1').column_name.should == 'B'
    end
    it 'は無効なセル名を指定したときnilとなります。' do
      CellName.new('A').column_name.should be_nil
    end
  end
  describe '#column_num' do
    it 'で列番号を取得できます' do
      CellName.new('A1').column_num.should == 1
      CellName.new('B1').column_num.should == 2
    end
    it 'は無効なセル名を指定したときnilとなります。' do
      CellName.new('A').column_num.should be_nil
    end
  end
  describe '#row_num' do
    it 'で行番号を取得できます' do
      CellName.new('A1').row_num.should == 1
      CellName.new('A2').row_num.should == 2
    end
    it 'は無効なセル名を指定したときnilとなります。' do
      CellName.new('1').row_num.should be_nil
    end
  end
  describe '#valid?' do
    it 'は有効なセルではtrueとなる' do
      CellName.new('A1').valid?.should be_truthy
    end
    it 'は無効なセルではfalseとなる' do
      CellName.new('A').valid?.should be_falsey
      CellName.new('1').valid?.should be_falsey
    end
    describe 'は列の範囲を判定して' do
      it 'は一文字の列名称の列を有効とします。' do
        CellName.new('A1').should be_valid
        CellName.new('Z1').should be_valid
      end
      it 'は2文字の列名称の列を有効とします。' do
        CellName.new('AA1').should be_valid
        CellName.new('ZZ1').should be_valid
      end
      it 'は3文字の列名称の列を有効とします。' do
        CellName.new('AAA1').should be_valid
        CellName.new('XFD1').should be_valid
      end
      it 'は大きすぎる列名を指定した場合は無効とします。' do
        CellName.new('XFE1').should_not be_valid
        CellName.new('AAAA1').should_not be_valid
        CellName.new('ZZZZ1').should_not be_valid
      end
    end
    describe 'は列の範囲を判定して' do
      it 'は存在する大きな行番号の行を有効とします。' do
        CellName.new('A1048576').should be_valid
        
      end
      it 'は大きすぎる行番号を指定した場合にnilを返します。' do
        CellName.new('A1048577').should_not be_valid
      end
    end
  end
  describe '#to_s' do
    it 'はA1形式で文字列を返します。' do
      CellName.new('A1').to_s.should == 'A1'
      CellName.new('B2').to_s.should == 'B2'
    end
  end
end