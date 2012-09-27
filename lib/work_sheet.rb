# coding: UTF-8
require 'rexml/document'

require 'cell'
require 'blank_value'
require 'cell_range'

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

  def cell_value(ref)
    cell(ref).value if cell(ref)
  end

  def cell(ref)
    ref = ref.to_s
    @cell_cache[ref] ||=
      if cell_xml(ref)
         Cell.new(cell_xml(ref), self, @tag_in_book)
      elsif is_ref(ref)
        Cell.new(ref, self, @tag_in_book)
      end
  end
  
  def range(*corner)
    case corner.size
    when 1
      corner_name1, corner_name2 = corner.first.to_s.split /:|_/
    when 2
      corner_name1, corner_name2 = *corner
    else
      raise ArgumentError, "wrong number of arguments (#{corner.size} for 1..2)"
    end

    return nil if !is_ref(corner_name1) || !is_ref(corner_name2)

    corner_names = upper_left_and_lower_right(corner_name1, corner_name2)

    @range_cache[corner_names] ||=
      CellRange.new(*corner_names, self)
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

private
  
  CELL_REGEXP = /^([A-Z]+)(\d+)$/
  
  def is_ref(ref)
    ref =~ CELL_REGEXP && [*'A'..$1].size <= 2**14 && $2.to_i <= 2**20
  end
  
  def ref_split(ref)
    ref.to_s =~ CELL_REGEXP
    [$1, $2]
  end

  def upper_left_and_lower_right(corner_name1, corner_name2)
    column_name1, row_name1 = ref_split(corner_name1)
    column_name2, row_name2 = ref_split(corner_name2)
    
    column_names = [column_name1, column_name2]
    row_names = [row_name1, row_name2]
    
    upper_left = column_names.min + row_names.min
    lower_right = column_names.max + row_names.max
    
    [upper_left, lower_right]
  end

  def cell_xml(ref)
    @xml.elements["./sheetData/row/c[@r='#{ref.to_s}']"]
  end

  def method_missing(method_name, *args)
    case method_name
    when /.*(?=\=$)/
      cell($&).value = args.first
    when /^[A-Z]\d_[A-Z]\d$/
      range(method_name)
    when /^[A-Z]\d$/
      value = cell_value(method_name)
      value.nil? ? super : value
    else
      super
    end
  end
end