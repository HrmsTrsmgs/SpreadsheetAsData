# coding:Shift_JIS

require 'spec_helper'

describe WorkBook, 'がShift_JISのコードから開かれた場合' do
  subject{WorkBook.open('./spec/test_data/Book1.xlsx')}

  after do
    subject.close
  end

  describe '#open' do
    context 'がShift_JISで書かれたコードから呼び出されてファイルを開く時ににエンコードを明示的に指定しなくても、' do
      it 'シート名を取得する文字コードがShift_JISになります。' do
        subject.sheets[2].name.should == 'いろいろなデータ'
      end

      it 'シートを選択する文字列は制限されません。'  do
        subject['いろいろなデータ'].should equal subject.sheets[2]
      end

      it 'シートを取得するメソッドは制限されません。'  do
        subject.いろいろなデータ.should equal subject.sheets[2]
      end

      it '文字列データの文字コードがShift_JISになります。' do
        subject.いろいろなデータ.A3.should == 'あいうえお'
      end
    end
  end
end