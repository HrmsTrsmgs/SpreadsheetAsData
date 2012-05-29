# coding: UTF-8

require 'rexml/document'

class Cell

  attr_reader :sheet

  # インスタンスを初期化します。
  def initialize(xml, sheet)
    if xml =~ /[A-Z]+[0-9]+/
      @ref = xml
    else
      @xml = xml
    end
    @sheet = sheet
  end

  def book
    sheet.book
  end

  def value
    @value ||=
      if not @xml
        BlankValue.new
      else
        case @xml.attributes['t']
        when nil
          @xml.elements[1].text.to_f
        when 'b'
          @xml.elements[1].text != '0'
        when 's'
          book.shared_strings[@xml.elements[1].text.to_i].encode(book.encoding)
        end
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