# coding:UTF-8

require 'fileutils'

require 'spec_helper'

describe WorkBook do

  BOOK_OPEND_EACH_TIME = test_file('逐次開かれるBook')

  subject{ TestFile.book1 }
  let(:book2){ TestFile.book2 }
  let(:anomaly){ TestFile.変則リレーション }
  let(:utf_8){WorkBook.open(TestFile.utf_8で開くbook_path, 'UTF-8')}
  let(:euc_jp){WorkBook.open(TestFile.euc_jpで開くbook_path, 'EUC-JP')}
  
  after do
    TestFile.close
  end
  
  after do
    subject.close
    book2.close
    anomaly.close
    utf_8.close
    euc_jp.close
  end

  describe '::open' do
    it 'の戻り値はWorkBookである。' do
      subject.class.should == WorkBook
    end

    it 'のブロック引数はWorkBookである。' do
      WorkBook.open(BOOK_OPEND_EACH_TIME) do |book|
        book.class.should == WorkBook
      end
    end

    context 'はファイル内のリレーション' do
      it 'が変則的な場合でも動作する。' do
        anomaly.Sheet1.C3.should == 4
      end
    end

    it 'でエンコードを指定することにより、シート名を取得する文字コードを選択できる。' do
      utf_8.sheets[2].name.should == 'いろいろなデータ'.encode('UTF-8')
      euc_jp.sheets[2].name.should == 'いろいろなデータ'.encode('EUC-JP')
    end

    it 'でエンコードを指定した時に、シートを選択する文字列のエンコードは制限されない。' do
      utf_8['いろいろなデータ'].should equal utf_8.sheets[2]
      euc_jp['いろいろなデータ'].should equal euc_jp.sheets[2]
    end

    it 'でエンコードを指定した時に、メソッド呼び出しのエンコードは制限されない。' do
      utf_8.いろいろなデータ.should equal utf_8.sheets[2]
      euc_jp.いろいろなデータ.should equal euc_jp.sheets[2]
    end

    it 'でエンコードを指定することで、文字列データを取得する文字コードを指定できる。' do
      utf_8.いろいろなデータ.A3.should == 'あいうえお'.encode('UTF-8')
      euc_jp.いろいろなデータ.A3.should == 'あいうえお'.encode('EUC-JP')
    end
  end

  describe '#close' do
    it 'の時に変更は保存されている。' do
      book = WorkBook.open(TestFile.book1_copy_path) do |book|
        book.Sheet1.cell(:A1).value = 999
      end
      
      WorkBook.open(TestFile.book1_copy_path) do |book|
        book.Sheet1.A1.should == 999
      end
    end
  end

  describe '#file_path' do
    it 'で指定したパス取得できる。' do
      subject.file_path.should == test_file('Book1')
      book2.file_path.should == test_file('Book2')
    end
    it 'が、open時にスラッシュ区切りでパスを指定した場合にも、指定した通りにパス取得できる。' do
      test_file_slash = BOOK_OPEND_EACH_TIME.gsub('\\', '/')
      WorkBook.open(test_file_slash) do |book|
        book.file_path.should == test_file_slash
      end
    end
  end

  describe '#sheets' do
    it 'でシートが取得できる。' do
      book2.sheets.map(&:name).should == %w[Sheet1 Sheet2]
    end
  end

  describe '#[]' do
    context 'の引数として文字列を指定した時。' do
      it 'にシートが取得できる' do
        subject['Sheet1'].should equal subject.sheets[0]
        subject['Sheet2'].should equal subject.sheets[1]
      end

      it 'に存在しない名前を指定された時にはnilを返す。' do
        subject['Sheet999'].should be_nil
      end

      it 'に同一シートを同一オブジェクトとして扱う。' do
        subject['Sheet1'].should equal subject['Sheet1']
      end

      it 'に日本語名のシートが取得できる。' do
        subject['いろいろなデータ'].should equal subject.sheets[2]
      end
    end
    context 'の引数としてシンボルを指定した時' do
      it 'にシートが取得できる。' do
        subject[:Sheet1].should equal subject.sheets[0]
        subject[:Sheet2].should equal subject.sheets[1]
      end
    end
  end

  describe '#シート名' do
    it 'がシートを取得できる。' do
      subject.Sheet1.should equal subject.sheets[0]
      subject.Sheet2.should equal subject.sheets[1]
    end

    it 'は存在しないシート名を指定された時にはNoMethodErrorを返す。' do
      ->{subject.Sheet999}.should raise_error NoMethodError
    end

    it 'で日本語名のシートが取得できる' do
      subject.いろいろなデータ.should equal subject.sheets[2]
    end
  end
end