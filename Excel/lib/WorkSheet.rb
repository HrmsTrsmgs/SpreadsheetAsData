require 'fileutils'
require 'rexml/document'

require 'Package'

class Cell

	# インスタンスを初期化します。
	def initialize(xml)
		@xml = xml
	end
	def value
		return @xml.elements[1].text.to_i
	end
end

# Excelのシートです。
class WorkSheet
	
	# インスタンスを初期化します。
	def initialize(tag, doc)
		@tag = tag
		@xml = doc
	end
	
	# シート名を取得します。
	def name
		return @tag.attributes['name'].encode('Shift_JIS')
	end
	
	def cell_value(ref)
		cell = Cell.new(@xml.elements.to_a('//c').find{|c| c.attributes['r'] == ref.to_s})
		return cell.value
	end
	
		def method_missing(method_name)
		self.cell_value(method_name) || super
	end
end