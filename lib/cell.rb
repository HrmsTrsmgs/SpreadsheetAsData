# coding: UTF-8

require 'rexml/document'

class Cell

  attr_reader :sheet

  # インスタンスを初期化します。
  def initialize(xml_or_ref, sheet, part)
    if xml_or_ref =~ /[A-Z]+[0-9]+/
      @ref = xml_or_ref
    else
      @xml = xml_or_ref
    end
    @sheet = sheet
    @part = part
  end

  def book
    sheet.book
  end

  def value
    if @value.nil?
      @value =
        if @xml
          text = @xml.elements['./v'].text
          case @xml.attributes['t']
          when nil
            text.to_f
          when 'b'
            text != '0'
          when 's'
            book.shared_strings[text.to_i].encode(book.encoding)
          end
        else
          BlankValue.new
        end
    end
    @value
  end

  def value=(value)
    @part.change
    @value = value
    @xml ||= sheet.add_cell_xml(@ref)
    v = @xml.elements['./v']
    case value
      when Numeric
        v.text = value.to_s
        @xml.attributes['t'] = nil
      when true, false
        v.text = value ? 1 : 0
        @xml.attributes['t'] = 'b'
      else
        v.text = book.shared_strings << value
        @xml.attributes['t'] = 's'
    end
    
  end
  
  def row_num
    /^[A-Z]+(\d+)$/ =~ to_s
    $1.to_i
  end

  def to_s
    if @xml
      @xml.attributes['r']
    else
      @ref
    end
  end

  alias ref to_s
end