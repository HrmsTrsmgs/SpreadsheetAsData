# coding: UTF-8

require 'rexml/document'

class PackagePart

  attr_reader :part_uri

  def initialize(package, part_uri)
    @package = package
    package.initialized_parts << self
    @part_uri = part_uri.gsub(/^\.\/?/, '')
    @cache = {}
  end
  
  def rels_uri
    File.join(File.dirname(part_uri), '_rels', File.basename(part_uri) + '.rels').gsub(/^\.\/?/, '')
  end
  
  def changed?
    @changed
  end

  def change
    @changed = true
  end

  def xml_document
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
      @cache[tag] ||= PackagePart.new(@package, File.join(File.dirname(@part_uri), tag.attributes['Target']))
  end

private
  def relation_tags
    @relation_tags ||= @package.relation_tags(self)
  end
end