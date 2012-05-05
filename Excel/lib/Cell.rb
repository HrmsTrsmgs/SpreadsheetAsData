require 'rexml/document'

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