# -*- encoding: utf-8 -*-
require File.dirname(__FILE__) + '/../lib/WorkBook'
describe WorkBook do
	def test_file(file_name)
		return '.\spec\test_data\\' + file_name + '.xlsx'
	end
	
	def unopened_file_name
		return '逐次開かれるBook'
	end
	
	after(:all) do
		subject.close
		book2.close
		anomaly.close
	end
	
	subject{WorkBook.open(test_file('Book1'))}
	let(:book2){WorkBook.open(test_file('Book2'))}
	let(:anomaly){WorkBook.open(test_file('変則リレーション'))}
	
	describe '#open' do
	
		it 'の戻り値はnilではない。' do
			should_not be_nil
		end
		
		it 'の戻り値はWorkBookである。' do
			subject.class.should == WorkBook
		end
	
		it 'のブロック引数はnilではない。' do
			WorkBook.open(test_file(unopened_file_name)) do |book|
				book.should_not be_nil
			end
		end
		
		it 'のブロック引数はWorkBookである。' do
			WorkBook.open(test_file(unopened_file_name)) do |book|
				book.class.should == WorkBook
			end
		end
		
		context 'ファイル内のリレーション' do
			it 'が変則的な場合でも動作する。' do
				anomaly.Sheet1.C3.should == 4
			end
		end
		context '内部ファイル操作' do
			it 'の時に解凍を行っている。' do
				temp_dir_name = test_file('tmp_' + unopened_file_name)
				delete_all(temp_dir_name) if Dir.exist?(temp_dir_name)
				WorkBook.open(test_file(unopened_file_name)) do |book|
					Dir.exist?(temp_dir_name).should == true
				end
			end
			
			it 'の終了時に解凍した作業ファイルの削除を行っている。' do
				temp_dir_name = test_file('tmp_' + unopened_file_name)
				delete_all(temp_dir_name) if Dir.exist?(temp_dir_name)
				WorkBook.open(test_file(unopened_file_name)) do |book|
				end
				Dir.exist?(temp_dir_name).should == false
			end
		end
	end
	describe '#file_path' do
		it 'が指定したパス取得できる。' do
			subject.file_path.should == test_file('Book1')
			book2.file_path.should == test_file('Book2')
		end
		it 'がスラッシュ区切りで指定した場合にも指定した通りにパス取得できる。' do
			test_file_slash = test_file(unopened_file_name).gsub('\\', '/')
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
		context '引数として文字列を指定' do
			it 'でシートが取得できる' do
				subject['Sheet1'].should equal subject.sheets[0]
				subject['Sheet2'].should equal subject.sheets[1]
			end
			
			it 'は存在しない名前を指定された時にはnilを返す。' do
				subject['Sheet999'].should be_nil
			end
			
			it 'は同一シートを同一オブジェクトとして扱う。' do
				subject['Sheet1'].should equal subject['Sheet1']
			end
			
			it 'で日本語名のシートが取得できる' do
				subject['いろいろなデータ'].should equal subject.sheets[2]
			end
		end
		context '引数としてシンボルを指定' do
			it 'でシートが取得できる' do
				subject[:Sheet1].should equal subject.sheets[0]
				subject[:Sheet2].should equal subject.sheets[1]
			end
		end
	end
	
	
	describe '#シート名' do
		it 'でシートが取得できる' do
			subject.Sheet1.should equal subject.sheets[0]
			subject.Sheet2.should equal subject.sheets[1]
		end
		
		it 'は存在しないシート名を指定された時にはNoMethodErrorを返す。' do
			lambda{subject.Sheet999}.should raise_error NoMethodError
		end
		
		it 'で日本語名のシートが取得できる' do
			subject.いろいろなデータ.should equal subject.sheets[2]
		end
	end
end

# 指定したパスのフォルダ及び、その下にあるファイル、フォルダをすべて削除します。
def delete_all(deleted)
	if FileTest.directory?(deleted) then
		Dir.foreach(deleted) do |child_path|
			next if child_path =~ /^\.+$/
			delete_all(deleted.sub(/\/+$/,"") + "/" + child_path )
		end
		Dir.rmdir(deleted) rescue ""
	else
		File.delete(deleted)
	end
end	