# -*- encoding: utf-8 -*-
require File.dirname(__FILE__) + '/../lib/WorkBook'
describe WorkBook do
	describe '#initialize' do
		it 'が変則的に' do
			WorkBook.open('.\spec\変則リレーション.xlsx') do |book|
				book.Sheet1.cell('C3').value.should == 4
			end
		end
	end
	describe '#open' do
		it 'が作業用に解凍を' do
			temp_dir_name = '.\spec\tmp_Book1.xlsx'
			delete_all(temp_dir_name) if Dir.exist?(temp_dir_name)
			WorkBook.open('.\spec\Book1.xlsx') do |book|
				Dir.exist?(temp_dir_name).should == true
			end
		end
		
		it 'が終了時に作業ファイルの削除を' do
			temp_dir_name = 'tmp_Book1.xlsx'
			delete_all(temp_dir_name) if Dir.exist?(temp_dir_name)
			WorkBook.open('.\spec\Book1.xlsx') do |book|
			end
			Dir.exist?(temp_dir_name).should == false
		end
		
		it 'の実行時にWorkBookオブジェクトの確認を' do
			WorkBook.open('.\spec\Book1.xlsx') do |book|
				book.should_not be_nil
			end
		end
		
		it 'の実行時にWorkBookオブジェクトの型がWorkBookであることの確認を' do
			WorkBook.open('.\spec\Book1.xlsx') do |book|
				book.class.should == WorkBook
			end
		end
	end
	describe '#file_path' do
		it '指定したパスが取得できることの確認を' do
			WorkBook.open('.\spec\Book1.xlsx') do |book|
				book.file_path.should == '.\spec\Book1.xlsx'
			end
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