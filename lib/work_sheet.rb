# coding: UTF-8
require 'rexml/document'

require 'cell'
require 'blank_value'

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
    @cell_hash[ref] ||=
      if cell_xml(ref)
        @cell_hash[ref] = Cell.new(cell_xml(ref), self)
      elsif ref =~ /[A-Z]+\d+/
        @cell_hash[ref] = Cell.new(ref, self)
      end
  end

  def method_missing(method_name, *args)
    value = self.cell_value(method_name)
    if value.nil?
      super
    else
      value
    end
  end

private
  def cell_xml(ref)
    @xml.elements.to_a('//c').find{|c| c.attributes['r'] == ref.to_s}
  end
end