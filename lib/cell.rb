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
    @value ||=
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

  def value=(value)
    @part.change
    @value = value
    @xml.elements['//v'].text = value.to_s
  end

  def ref
    if @xml
      @xml.attributes['r']
    else
      @ref
    end
  end
end