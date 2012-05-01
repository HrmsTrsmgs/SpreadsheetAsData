# -*- encoding: utf-8 -*-
require 'test/unit'
require File.dirname(__FILE__) +  '/lib/WorkBook'

# WorkBookクラスのテストを行います。
class WorkBookTest < Test::Unit::TestCase

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
	
	# file_pathアクセサで指定したパスが取得できることを確認します。
	def test_file_path
		WorkBook.open('.\spec\Book1.xlsx') do |book|
			assert_equal('.\spec\Book1.xlsx', book.file_path, 'Book1のテストでfile_pathが正しくありません')
		end
		WorkBook.open('.\spec\Book2.xlsx') do |book|
			assert_equal('.\spec\Book2.xlsx', book.file_path, 'Book2のテストでfile_pathが正しくありません')
		end
	end
	
	# パスをスラッシュ区切りで指定した場合にもfile_pathアクセサで指定したパスが取得できることを確認します。
	def test_file_path_path_as_slash
		WorkBook.open('./spec/Book1.xlsx') do |book|
			assert_equal('./spec/Book1.xlsx', book.file_path, 'Book1のテストでfile_pathが正しくありません')
		end
		WorkBook.open('./spec/Book1.xlsx') do |book|
			assert_equal('./spec/Book1.xlsx', book.file_path, 'Book2のテストでfile_pathが正しくありません')
		end
	end
	
	# sheetsアクセサでシートが取得できることを確認します。
	def test_sheets
		WorkBook.open('.\spec\Book1.xlsx') do |book|
			assert_equal(3, book.sheets.size, 'Book1のテストでシート数が適切でありません。')
		end
		WorkBook.open('.\spec\Book2.xlsx') do |book|
			assert_equal(2, book.sheets.size, 'Book2のテストでシート数が適切でありません')
		end
	end
	
	# シートが[]で呼び出せることを確認します。
	def test_bracked
		WorkBook.open('.\spec\Book1.xlsx') do |book|
			assert_equal('Sheet1', book['Sheet1'].name, 'Sheet1が取得できませんでした')
			assert_equal('Sheet2', book['Sheet2'].name, 'Sheet2が取得できませんでした')
		end
	end
	
	# []で存在しないシートを呼び出そうとした場合にnilを返すことを確認します。
	def test_bracket_not_exist_name
		WorkBook.open('.\spec\Book1.xlsx') do |book|
			assert_nil(book['Sheet999'], '存在しないシートが取得されました')
		end
	end
	
	# メソッド名としてシートが呼び出されることを確認します。
	def test_sheet_name_as_method
		WorkBook.open('.\spec\Book1.xlsx') do |book|
			assert_equal('Sheet1', book.Sheet1.name, 'Sheet1が取得できませんでした')
			assert_equal('Sheet2', book.Sheet2.name, 'Sheet2が取得できませんでした')
		end
	end
	
	# メソッド名で存在しないシートを呼び出そうとした場合にnilを返すことを確認します。
	def test_sheet_name_as_method_not_exist_name
		WorkBook.open('.\spec\Book1.xlsx') do |book|
			assert_raise(NoMethodError, '存在しないシートの指定でエラーが出ませんでした'){book.Sheet999}
		end
	end
	
	# 同一シートが同一オブジェクトとして取得できることを確認します。
	def test_same_sheet
		WorkBook.open('.\spec\Book1.xlsx') do |book|
			assert_same(book.Sheet1, book.Sheet1, '一つのシートに二つオブジェクトがあります')
		end
	end
end