# coding: UTF-8

require 'rexml/document'

class SharedStrings
  # インスタンスを初期化します。
  def initialize(part)
    @part = part
  end
  
  def <<(string)
    @part.xml_document.elements['//sst'].add_element(new_string_tag(string))
    @part.change
    @part.xml_document.elements['//sst'].elements.size - 1
  end
  
  def [](index)
    @part.xml_document.elements.to_a('//t').map {|t| t.text }[index]
  end
  
private
  def new_string_tag(string)
    t = REXML::Element.new('t')
    t.text = string
    phonetic = REXML::Element.new('phoneticPr')
    phonetic.attributes['fontId'] = 1
    si = REXML::Element.new('si')
    si.add_element(t)
    si.add_element(phonetic)
    si
  end
end