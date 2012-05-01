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
		return @tag.attributes['name']
	end
	
	def cell(ref)
		return Cell.new(@xml.elements.to_a('//c').select{|c| c.attributes['r'] == ref}[0])
	end
end