# coding: UTF-8
require 'rexml/document'

require 'cell'
require 'blank_value'

# Excelのシートです。
class WorkSheet

  attr_reader :book

  # インスタンスを初期化します。
  def initialize(tag, doc, book, doc_part)
    @tag = tag
    @xml = doc
    @book = book
    @doc_part = doc_part
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
        @cell_hash[ref] = Cell.new(cell_xml(ref), self, @doc_part)
      elsif ref =~ /[A-Z]+\d+/
        @cell_hash[ref] = Cell.new(ref, self, @doc_part)
      end
  end
  
  def add_cell_xml(ref)
      v = REXML::Element.new('v')
      c = REXML::Element.new('c')
      c.attributes['r'] = ref
      c.add_element(v)
      ref =~ /\d+/
      row = @xml.get_elements('//row').find {|row| row.attributes['r'] == $& }
      unless row
        row = REXML::Element.new('row')
        row.attributes['r'] = $&
        @xml.elements['//sheetData'].add_element(row)
      end
      row.add_element(c)
      @xml.get_elements('//c').find{|c| c.attributes['r'] == ref.to_s}
  end

  def method_missing(method_name, *args)
    if method_name =~ /.*(?=\=$)/
      cell($&).value = args.first
    else
      value = self.cell_value(method_name)
      if value.nil?
        super
      else
        value
      end
    end
  end

private
  def cell_xml(ref)
    @xml.elements.to_a('//c').find{|c| c.attributes['r'] == ref.to_s}
  end
end