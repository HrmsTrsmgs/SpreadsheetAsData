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
          case @xml.attributes['t']
          when nil
            @xml.elements[1].text.to_f
          when 'b'
            @xml.elements[1].text != '0'
          when 's'
            book.shared_strings[@xml.elements[1].text.to_i].encode(book.encoding)
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
    if not @xml
      v = REXML::Element.new('v')
      v.text = value.to_s
      c = REXML::Element.new('c')
      c.attributes['r'] = @ref
      c.add_element(v)
      @ref =~ /\d+/
      row = sheet.xml.get_elements('//row').find {|row| row.attributes['r'] == $& }
      row.add_element(c)
      @xml = sheet.xml.elements.to_a('//c').find{|c| c.attributes['r'] == @ref.to_s}
    end
    
    case value
      when Numeric
        @xml.elements['//v'].text = value.to_s
        @xml.attributes['t'] = nil
      when true, false
        @xml.elements['//v'].text = value ? 1 : 0
        @xml.attributes['t'] = 'b'
      else
        @xml.elements['//v'].text = book.shared_strings << value
        @xml.attributes['t'] = 's'
    end
  end

  def ref
    if @xml
      @xml.attributes['r']
    else
      @ref
    end
  end
end