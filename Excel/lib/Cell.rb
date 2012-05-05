require 'rexml/document'

class Cell
	
	attr_reader :sheet
		
	# インスタンスを初期化します。
	def initialize(xml, sheet)
		@xml = xml
		@sheet = sheet
	end
	
	def book
		sheet.book
	end
	
	def value
		case @xml.attributes['t']
		when nil
			@xml.elements[1].text.to_i
		when 'b'
			@xml.elements[1].text != '0'
		when 's'
			book.shared_strings[@xml.elements[1].text.to_i].encode('Shift_JIS')
		end
	end
end