require 'rexml/document'

class Cell
	
	attr_reader :sheet
		
	# インスタンスを初期化します。
	def initialize(xml, sheet)
		if xml =~ /[A-Z]+[0-9]+/ then
			@ref = xml
		else
			@xml = xml
		end
		@sheet = sheet
	end
	
	def book
		sheet.book
	end
	
	def value
	
		return @value = BlankValue.new until @xml
		@value ||=
			case @xml.attributes['t']
			when nil
				@xml.elements[1].text.to_f
			when 'b'
				@xml.elements[1].text != '0'
			when 's'
				book.shared_strings[@xml.elements[1].text.to_i].encode(book.encoding)
			end
	end
	
	def ref
		if @xml then
			@xml.attributes['r']
		else
			@ref
		end
	end
end