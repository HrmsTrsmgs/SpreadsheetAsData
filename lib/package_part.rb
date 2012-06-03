# coding: UTF-8
require 'rexml/document'

class PackagePart

  attr_reader :part_uri

  def initialize(package, part_uri)
    @package = package
    package.initialized_parts << self
    @part_uri = part_uri
    @cache = {}
  end
  
  def changed?
    @changed
  end

  def change
    @changed = true
  end

  # ブック情報を記述してあるWorkBook.xmlドキュメントを取得します。
  def xml_document
    #workbook.xmlのパスは変更するとExcelでも起動できなくなるため、変更には対応しません。
    @xml_document ||= @package.xml_document(self)
  end

  def relation(key)
    tag =
      if key.respond_to?(:attributes) && key.attributes['r:id']
        relation_tags.find {|tag| tag.attributes['Id'] == key.attributes['r:id'] }
      elsif key =~ %r!^http://schemas.openxmlformats.org/officeDocument/2006/relationships/!
        relation_tags.find {|tag| tag.attributes['Type'] == key }
      else
        relation_tags.find {|tag| tag.attributes['Id'] == key }
      end
      
    tag && 
      @cache[tag] ||= PackagePart.new(@package, File.dirname(@part_uri) +"/" + tag.attributes['Target'])
  end

private
  def relation_tags
    @relation_tags ||= @package.relation_tags(self)
  end
end