# coding: UTF-8
require 'rexml/document'

require_relative 'cell'
require_relative 'blank_value'
require_relative 'cell_range'
require_relative 'cell_name'

# Excelのシートです。
class WorkSheet

  attr_reader :book

  # インスタンスを初期化します。
  def initialize(tag, doc, book, tag_in_book)
    @tag = tag
    @xml = doc.root
    @book = book
    @tag_in_book = tag_in_book
    @cell_cache = {}
    @range_cache = {}
  end

  # シート名を取得します。
  def name
    @tag.attributes['name'].encode(book.encoding)
  end

  def cell_value(cell_name)
    cell(cell_name).value if cell(cell_name)
  end

  def blank?(cell_name)
    cell_value(cell_name).class == BlankValue
  end

  def cell(ref)
    ref = ref.to_s
    @cell_cache[ref] ||=
      if cell_xml(ref)
        Cell.new(cell_xml(ref), self, @tag_in_book)
      elsif CellName.valid?(ref)
        Cell.new(ref, self, @tag_in_book)
      end
  end
  
  def defined_name
    sheet_tables
  end
  
  def sheet_tables
    Array.new(@xml.get_elements('//tableParts/tablePart').size)
  end
  
  def range(*corner)
    raise ArgumentError, "wrong number of arguments (#{corner.size} for 1..2)" if 3 <= corner.size

    name = RangeName.new(*corner)
    
    @range_cache[name] ||=
      name.valid? ? CellRange.new(name.upper_left, name.lower_right, self) : nil
  end

  def add_cell_xml(ref)
      v = REXML::Element.new('v')
      c = REXML::Element.new('c')
      c.attributes['r'] = ref
      c.add_element(v)
      ref =~ /\d+/
      row = rows_xml.find {|row| row.attributes['r'] == $& }
      unless row
        row = REXML::Element.new('row')
        row.attributes['r'] = $&
        @xml.elements['./sheetData'].add_element(row)
      end
      row.add_element(c)
  end
  
  def real_row_nums
    rows_xml.map{|row| row.attributes['r'].to_i  }
  end
private

  def rows_xml
    @xml.get_elements('./sheetData/row')
  end
  def ref_split(ref)
    name = CellName.new(ref)
    [name.column_name, name.row_num]
  end

  def cell_xml(ref)
    @xml.elements["./sheetData/row/c[@r='#{ref.to_s}']"]
  end

  def method_missing(method_name, *args)
    case method_name
    when /.*(?=\=$)/
      cell($&).value = args.first
    when RangeName.valid?
      range(method_name)
    when CellName.valid?
      cell_value(method_name)
    else
      super
    end
  end
end