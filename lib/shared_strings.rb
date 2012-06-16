# coding: UTF-8

require 'rexml/document'

class SharedStrings
  # インスタンスを初期化します。
  def initialize(part)
    @part = part
  end
  
  def <<(string)
    strings.find_index(string) ||
      begin
        sst_tag.add_element(new_t(string))
        @part.change
        size - 1
      end
  end
  
  def [](index)
    strings[index]
  end
  
  def size
    sst_tag.elements.size
  end
  
private
  def sst_tag
    @part.xml_document.elements['//sst']
  end
  
  def t_tags
    @part.xml_document.get_elements('//t')
  end
  
  def strings
    t_tags.map {|t| t.text }
  end
  
  def new_t(string)
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