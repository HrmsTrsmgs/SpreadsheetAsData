# coding: UTF-8
require 'rexml/document'

require 'cell'
require 'blank_value'
require 'cell_range'
require 'cell_name'

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
  
  def defined_name
    Array.new(@xml.get_elements('//tableParts/tablePart').size)
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
    CellName.new(ref).valid?
  end
  
  def ref_split(ref)
    name = CellName.new(ref)
    [name.column_name, name.row_num]
  end

  def upper_left_and_lower_right(corner_name1, corner_name2)
    column_name1, row_num1 = ref_split(corner_name1)
    column_name2, row_num2 = ref_split(corner_name2)
    
    column_names = [column_name1, column_name2]
    row_nums = [row_num1, row_num2]
    
    upper_left = column_names.min + row_nums.min.to_s
    lower_right = column_names.max + row_nums.max.to_s
    
    [upper_left, lower_right]
  end

  def cell_xml(ref)
    @xml.elements["./sheetData/row/c[@r='#{ref.to_s}']"]
  end

  def method_missing(method_name, *args)
    case method_name
    when /.*(?=\=$)/
      cell($&).value = args.first
    when /^[A-Z]+\d+_[A-Z]+\d+$/
      value = range(method_name)
      value.nil? ? super : value
    when /^[A-Z]+\d+$/
      value = cell_value(method_name)
      value.nil? ? super : value
    else
      super
    end
  end
end