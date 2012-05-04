require 'fileutils'
require 'rexml/document'

require 'Package'

class Cell

	# インスタンスを初期化します。
	def initialize(xml)
		@xml = xml
	end
	def value
		case @xml.attributes['t']
		when nil
			@xml.elements[1].text.to_i
		when 'b'
			@xml.elements[1].text != '0'
		end
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
		p cell
		return cell.value
	end
	
	def method_missing(method_name)
		value = self.cell_value(method_name)
		if value.nil? then
			super
		else
			value
		end
	end
end