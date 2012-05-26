# -*- encoding: UTF-8 -*- 

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
    context 'がUTF-8で書かれたコードから呼び出されてファイルを開く時ににエンコードを明示的に指定しなくても、'.encode('Shift_JIS') do
      it 'シート名を取得する文字コードがUTF-8になる。'.encode('Shift_JIS') do
        subject.sheets[2].name.should == 'いろいろなデータ'
      end

      it 'シートを選択する文字列は制限されない。'.encode('Shift_JIS')  do
        subject['いろいろなデータ'].should equal subject.sheets[2]
      end

      it 'シートを取得するメソッドは制限されない。'.encode('Shift_JIS')  do
        subject.いろいろなデータ.should equal subject.sheets[2]
      end

      it '文字列データの文字コードがUTF-8になる。'.encode('Shift_JIS') do
        subject.いろいろなデータ.A3.should == 'あいうえお'
      end
    end
  end
end