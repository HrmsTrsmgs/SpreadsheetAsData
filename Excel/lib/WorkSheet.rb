require 'fileutils'
require 'rexml/document'

require 'Package'
require 'Cell'

# Excelのシートです。
class WorkSheet

	attr_reader :book

	# インスタンスを初期化します。
	def initialize(tag, doc, book)
		@tag = tag
		@xml = doc
		@book = book
		@cell_hash = {}
	end
	
	# シート名を取得します。
	def name
		@tag.attributes['name'].encode(book.encoding)
	end
	
	def cell_value(ref)
		cell(ref).value if cell(ref)
	end
	
	def cell(ref)
		cell = @cell_hash[ref]
		return cell if cell
		xml = @xml.elements.to_a('//c').find{|c| c.attributes['r'] == ref.to_s}
		if xml then
			cell = Cell.new(xml, self)
			@cell_hash[ref] = cell
		elsif ref =~ /[A-Z]+\d+/ then
			cell = Cell.new(ref, self)
			@cell_hash[ref] = cell
		end
	end
	
	def method_missing(method_name)
		value = self.cell_value(method_name)
		if value.nil? then super else value end
	end
end