require 'rexml/document'

class Cell
	# インスタンスを初期化します。
	def initialize(xml, sheet)
		@xml = xml
		@sheet = sheet
	end
	
	attr_reader :sheet
	
	def book
		sheet.book
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