# coding: UTF-8

require 'rexml/document'

class Cell

  attr_reader :sheet

  # インスタンスを初期化します。
  def initialize(xml, sheet, part)
    if xml =~ /[A-Z]+[0-9]+/
      @ref = xml
    else
      @xml = xml
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
    case value
      when Numeric
        @xml.elements['//v'].text = value.to_s
        @xml.attributes['t'] = nil
      when true, false
        @xml.elements['//v'].text = value ? 1 : 0
        @xml.attributes['t'] = 'b'
      else
        @xml.elements['//v'].text = book.set_shared_string(value)
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