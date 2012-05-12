# -*- encoding: UTF-8 -*- 
require File.dirname(__FILE__) + '/../lib/WorkBook'
describe WorkBook do
	def test_file(file_name)
		'.\spec\test_data\\' + file_name + '.xlsx'
	end
	
	after(:all) do
		subject.close
	end
	
	subject{WorkBook.open(test_file('Book1'))}
	
	describe '#open' do
		context 'UTF-8のファイルでファイルを開く'.encode('Shift_JIS') do
			it 'で指定せずにUTF-8のコードから読み出し、シート名を取得する文字コードをUTF-8で指定できる'.encode('Shift_JIS') do
				subject.sheets[2].name.should == 'いろいろなデータ'
			end
			
			it 'で指定せずにUTF-8のコードから読み出し、、シートを選択する文字列は制限されない。'.encode('Shift_JIS')  do
				subject['いろいろなデータ'].should equal subject.sheets[2]
			end
			
			it 'で指定せずにUTF-8のコードから読み出し、、シートを取得するメソッドは制限されない。'.encode('Shift_JIS')  do
				subject.いろいろなデータ.should equal subject.sheets[2]
			end
			
			it 'でエンコードを指定することで、文字列データを取得する文字コードを指定できる。'.encode('Shift_JIS') do
				subject.いろいろなデータ.A3.should == 'あいうえお'
			end
		end
	end
end	