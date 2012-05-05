$LOAD_PATH << File.dirname(__FILE__)

require 'zip/zipfilesystem'
require 'fileutils'
require 'rexml/document'

require 'WorkSheet'
require 'Package'

# Excelブックを扱うクラスです。
class WorkBook

	# 開いたExcelファイルのパスです。
	def file_path
		@package.file_path
	end
	
	# 新しいインスタンスの初期化を行います。
	def initialize(file_path)
		@package = Package.new(file_path)
		@part = @package.part('/xl/workbook.xml')
	end

	# Excelファイルのパスを指定し、開きます。<br/>
	# <br/>
	# ==== <b>[PARAM]file_path</b><br/>
	#  開くファイルのパスを指定します。
	def self.open(file_path)
			book = WorkBook.new(file_path)
			if block_given? then
				yield book
				book.close
			else
				return book
			end
	end
	
	# ファイルの操作を終了し、ファイルを開放します。
	def close
		@package.close
	end
	
	# ブックが持つシートを取得します。
	def sheets
		@sheets ||=
			@part.xml_document.elements.to_a('//sheet')
			.map{ |tag| WorkSheet.new(tag, @part.relations[tag].xml_document, self)}
	end
	
	# 名前を指定し、ブックが持つシートを取得します。
	def [](name)
		sheets.find{|sheet| sheet.name == name.to_s.encode('Shift_JIS')}
	end
	
	# 指定したメソッド名が定義されていない場合に呼び出されます。
	def method_missing(method_name)
		self[method_name] || super
	end
end