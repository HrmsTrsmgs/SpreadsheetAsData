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

  def relation_tags
    @relation_tags ||= @package.relation_tags(@part_uri)
  end

  def relation_cache
    if !@hash then
      relations_array =
        relation_tags
          .map{|tag| 
            [tag,
            PackagePart.new(
              @package,
              File.dirname(@part_uri) +"/" + tag.attributes['Target'])
            ]
          }
      @hash = Hash[*relations_array.flatten(1)]
    end

    @hash
  end

  def relation(key)
      if key.respond_to?(:attributes) && key.attributes['r:id'] then
        relation_cache[relation_tags.find {|tag| tag.attributes['Id'] == key.attributes['r:id'] }]
      elsif key =~ %r!^http://schemas.openxmlformats.org/officeDocument/2006/relationships/! then
        relation_cache[relation_tags.find {|tag| tag.attributes['Type'] == key }]
      else
        relation_cache[relation_tags.find {|tag| tag.attributes['Id'] == key }]
      end
  end

private
  def relation_tags
    @relation_tags ||= @package.relation_tags(self)
  end
end