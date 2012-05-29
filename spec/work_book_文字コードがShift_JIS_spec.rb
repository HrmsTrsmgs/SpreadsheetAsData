# coding:Shift_JIS

require 'spec_helper'

describe WorkBook do
  subject{WorkBook.open(test_file('Book1'))}

  def test_file(file_name)
    "./spec/test_data/#{file_name}.xlsx"
  end

  after(:all) do
    subject.close
  end

  describe '#open' do
    context 'がShift_JISで書かれたコードから呼び出されてファイルを開く時ににエンコードを明示的に指定しなくても、' do
      it 'シート名を取得する文字コードがShift_JISになる。' do
        subject.sheets[2].name.should == 'いろいろなデータ'
      end

      it 'シートを選択する文字列は制限されない。'  do
        subject['いろいろなデータ'].should equal subject.sheets[2]
      end

      it 'シートを取得するメソッドは制限されない。'  do
        subject.いろいろなデータ.should equal subject.sheets[2]
      end

      it '文字列データの文字コードがShift_JISになる。' do
        subject.いろいろなデータ.A3.should == 'あいうえお'
      end
    end
  end
end