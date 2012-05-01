# -*- encoding: utf-8 -*-
require 'test/unit'
require File.dirname(__FILE__) +  '/lib/WorkBook'

# WorkSheetクラスのテストを行います。
class WorkSheetTest < Test::Unit::TestCase

	# シートのリレーションが変則的に指定された
	def test_initialize_strange
		WorkBook.open('.\spec\変則リレーション.xlsx') do |book|
			assert_equal(4, book.Sheet1.cell('C3').value, 'A1セルの値が取得できません。')
		end
	end

	# nameアクセサーで取得できることを確認します。
	def test_name
		WorkBook.open('.\spec\Book1.xlsx') do |book|
			assert_equal('Sheet1', book.sheets[0].name, 'Sheet1のテストでnameプロパティが取得できていません')
			assert_equal('Sheet2', book.sheets[1].name, 'Sheet2のテストでnameプロパティが取得できていません')
		end
	end
	
	# cellメソッドで取得できることを確認します。
	def test_cell
		WorkBook.open('.\spec\Book1.xlsx') do |book|
			assert_equal(4, book.Sheet1.cell('C3').value, 'A1セルの値が取得できません。')
		end
	end
	
	
	# cellメソッドで行方向の取得が適切であることを確認します。
	def test_cell_row
		WorkBook.open('.\spec\Book1.xlsx') do |book|
			assert_equal(1, book.Sheet1.cell('A1').value, 'A1セルの値が取得できません。')
			assert_equal(2, book.Sheet1.cell('C1').value, 'C1セルの値が取得できません。')
		end
	end
	
	# cellメソッドで列方向の取得が適切であることを確認します。
	def test_cell_column
		WorkBook.open('.\spec\Book1.xlsx') do |book|
			assert_equal(1, book.Sheet1.cell('A1').value, 'A1セルの値が取得できません。')
			assert_equal(3, book.Sheet1.cell('A3').value, 'A3セルの値が取得できません。')
		end
	end
	
	# cellメソッドで各シートのセルが取得できることを確認します。
	def test_cell_sheet
		WorkBook.open('.\spec\Book1.xlsx') do |book|
			assert_equal(1, book.Sheet1.cell('A1').value, 'Sheet1のセルの値が取得できません。')
			assert_equal(5, book.Sheet2.cell('A1').value, 'Sheet2のセルの値が取得できません。')
		end
	end
	
end