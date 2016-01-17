require 'fileutils'

require 'spec_helper'

describe WorkBook do

  subject{ TestFile.book1 }
  let(:book2){ TestFile.book2 }
  let(:anomaly){ TestFile.変則リレーション }
  let(:utf_8){WorkBook.open(TestFile.book1_path, 'UTF-8')}
  let(:euc_jp){WorkBook.open(TestFile.book1_path, 'EUC-JP')}
  
  after do
    TestFile.close
  end

  describe '::open' do
    it 'の戻り値はWorkBookです。' do
      subject.class.should === WorkBook
    end

    it 'のブロック引数はWorkBookです。' do
      WorkBook.open(TestFile.book1_path) do |book|
        book.class.should == WorkBook
      end
    end
    
    it 'はファイルを束縛します。' do
      book = WorkBook.open(TestFile.book1_copy_path)
      expect { File.delete(TestFile.book1_copy_path) }.to raise_error Errno::EACCES
      book.close
    end
    
    it 'はブロック内でファイルを束縛します。' do
      WorkBook.open(TestFile.book1_copy_path) do |book|
        expect { File.delete(TestFile.book1_copy_path) }.to raise_error Errno::EACCES
      end
    end
    
    it 'はブロック終了時にファイルを開放します。' do
      WorkBook.open(TestFile.book1_copy_path) do |book|
      end
      expect { File.delete(TestFile.book1_copy_path) }.to_not raise_error
    end
    
    it 'は例外発生時もブロック終了時にファイルを開放します。' do
      begin
        WorkBook.open(TestFile.book1_copy_path) do |book|
          raise Exception
        end
      rescue Exception
      end
      expect { File.delete(TestFile.book1_copy_path) }.to_not raise_error
    end

    context 'はファイル内のリレーション' do
      it 'が変則的な場合でも動作します。' do
        anomaly.Sheet1.C3.should == 4
      end
    end

    it 'でエンコードを指定することにより、シート名を取得する文字コードを選択できます。' do
      utf_8.sheets[2].name.should == 'いろいろなデータ'.encode('UTF-8')
      euc_jp.sheets[2].name.should == 'いろいろなデータ'.encode('EUC-JP')
    end

    it 'でエンコードを指定した時に、シートを選択する文字列のエンコードは制限されません。' do
      utf_8['いろいろなデータ'].should equal utf_8.sheets[2]
      euc_jp['いろいろなデータ'].should equal euc_jp.sheets[2]
    end

    it 'でエンコードを指定した時に、メソッド呼び出しのエンコードは制限されません。' do
      utf_8.いろいろなデータ.should equal utf_8.sheets[2]
      euc_jp.いろいろなデータ.should equal euc_jp.sheets[2]
    end

    it 'でエンコードを指定することで、文字列データを取得する文字コードを指定できます。' do
      utf_8.いろいろなデータ.A3.should == 'あいうえお'.encode('UTF-8')
      euc_jp.いろいろなデータ.A3.should == 'あいうえお'.encode('EUC-JP')
    end
  end

  describe '#close' do
  
    it 'はファイルを開放します。' do
      book = WorkBook.open(TestFile.book1_copy_path)
      book.close
      expect { File.delete(TestFile.book1_copy_path) }.to_not raise_error
    end
  
    it 'の時に変更は保存されています。' do
      pending '書き込み機能は一時不可'
      book = WorkBook.open(TestFile.book1_copy_path) do |book|
        book.Sheet1.cell(:A1).value = 999
      end
      
      WorkBook.open(TestFile.book1_copy_path) do |book|
        book.Sheet1.A1.should == 999
      end
    end
  end

  describe '#file_path' do
    it 'で指定したパス取得できます。' do
      subject.file_path.should == test_file('Book1')
      book2.file_path.should == test_file('Book2')
    end
    it 'が、open時にスラッシュ区切りでパスを指定した場合にも、指定した通りにパス取得できます。' do
      test_file_slash = TestFile.book1_path.gsub('\\', '/')
      WorkBook.open(test_file_slash) do |book|
        book.file_path.should == test_file_slash
      end
    end
  end

  describe '#sheets' do
    it 'でシートが取得できます。' do
      book2.sheets.map(&:name).should == %w[Sheet1 Sheet2]
    end
  end

  describe '#[]' do
    context 'の引数として文字列を指定した時' do
      it 'にシートが取得できます。' do
        subject['Sheet1'].should equal subject.sheets[0]
        subject['Sheet2'].should equal subject.sheets[1]
      end

      it 'に存在しない名前を指定された時にはnilを返します。' do
        subject['Sheet999'].should be_nil
      end

      it 'に同一シートを同一オブジェクトとして扱います。' do
        subject['Sheet1'].should equal subject['Sheet1']
      end

      it 'に日本語名のシートが取得できます。' do
        subject['いろいろなデータ'].should equal subject.sheets[2]
      end
    end
    context 'の引数としてシンボルを指定した時' do
      it 'にシートが取得できます。' do
        subject[:Sheet1].should equal subject.sheets[0]
        subject[:Sheet2].should equal subject.sheets[1]
      end
    end
  end

  describe '#シート名' do
    it 'がシートを取得できます。' do
      subject.Sheet1.should equal subject.sheets[0]
      subject.Sheet2.should equal subject.sheets[1]
    end

    it 'は存在しないシート名を指定された時にはNoMethodErrorを返します。' do
      ->{subject.Sheet999}.should raise_error NoMethodError
    end

    it 'で日本語名のシートが取得できます。' do
      subject.いろいろなデータ.should equal subject.sheets[2]
    end
  end
  
  context 'に保存方法として' do
    before { skip }
    it 'cell.valueを使った書き込みは保存されています。' do
      book = WorkBook.open(TestFile.book1_copy_path) do |book|
        book.Sheet1.cell(:A1).value = 999
      end
      
      WorkBook.open(TestFile.book1_copy_path) do |book|
        book.Sheet1.A1.should == 999
      end
    end
    
    it 'セル名を使った書き込みは保存されています。' do
      book = WorkBook.open(TestFile.book1_copy_path) do |book|
        book.Sheet1.A1 = 999
      end
      
      WorkBook.open(TestFile.book1_copy_path) do |book|
        book.Sheet1.A1.should == 999
      end
    end
  end

  context 'に保存するデータとして' do
    before { skip }
    it '整数値の書き込みは保存されています。' do
      book = WorkBook.open(TestFile.book1_copy_path) do |book|
        book.Sheet1.A1 = 666
        book.Sheet1.C3 = 999
      end
      
      WorkBook.open(TestFile.book1_copy_path) do |book|
        book.Sheet1.A1.should == 666
        book.Sheet1.C3.should == 999
      end
    end
    
    it 'trueの書き込みは保存されています。' do
      book = WorkBook.open(TestFile.book1_copy_path) do |book|
        book.Sheet1.A1 = true
      end
      
      WorkBook.open(TestFile.book1_copy_path) do |book|
        book.Sheet1.A1.should == true
      end
    end
    
    it 'falseの書き込みは保存されています。' do
      book = WorkBook.open(TestFile.book1_copy_path) do |book|
        book.Sheet1.A1 = false
      end
      
      WorkBook.open(TestFile.book1_copy_path) do |book|
        book.Sheet1.A1.should == false
      end
    end
    
    it '文字列の書き込みは保存されています。' do
      book = WorkBook.open(TestFile.book1_copy_path) do |book|
        book.Sheet1.A1 = 'abcde'
        book.Sheet1.C3 = 'fghi'
      end
      
      WorkBook.open(TestFile.book1_copy_path) do |book|
        book.Sheet1.A1.should == 'abcde'
        book.Sheet1.C3.should == 'fghi'
      end
    end
  end
  
  it 'の空白セルに書き込みがされています。' do
    pending '書き込み機能は一時不可'
    book = WorkBook.open(TestFile.book1_copy_path) do |book|
      book.Sheet1.B1 = 999
    end
    
    WorkBook.open(TestFile.book1_copy_path) do |book|
      book.Sheet1.B1.should == 999
    end
  end

  it 'の空白行に書き込みがされています。' do
    pending '書き込み機能は一時不可'
    book = WorkBook.open(TestFile.book1_copy_path) do |book|
      book.Sheet1.B2 = 999
    end
    
    WorkBook.open(TestFile.book1_copy_path) do |book|
      book.Sheet1.B2.should == 999
    end
  end
  
  it 'の空白セルに書き込みがされた時に、行の整合性が取れています。' do
    pending '書き込み機能は一時不可'
    book = WorkBook.open(TestFile.book1_copy_path) do |book|
      book.Sheet1.B1 = 999
    end
    
    WorkBook.open(TestFile.book1_copy_path) do |book|
      book.Sheet1.B1.should == 999
      book.Sheet1.A1.should == 1
      book.Sheet1.real_row_nums.should == [1, 3]
    end
  end
end