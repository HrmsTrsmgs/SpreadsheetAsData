# coding: UTF-8

require 'rexml/document'

class PackagePart

  attr_reader :part_uri

  def initialize(package, part_uri)
    @package = package
    package.initialized_parts << self
    @part_uri = part_uri.gsub(/^\.\/?/, '')
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

  def relation_cache
    @relations_by_tag ||= relation_tags.each_with_object({}) do |tag, parts|
      uri = File.dirname(@part_uri) + "/" + tag.attributes['Target']
      parts[tag] = PackagePart.new(@package, uri)
    end
  end

  def relation(key)
    attribute, value = relation_lookup(key)
    relation_cache[relation_tags.find { |tag| tag.attributes[attribute] == value }]
  end

private
  def relation_lookup(key)
    if key.respond_to?(:attributes) && key.attributes['r:id']
      ['Id', key.attributes['r:id']]
    elsif key =~ %r!^http://schemas.openxmlformats.org/officeDocument/2006/relationships/!
      ['Type', key]
    else
      ['Id', key]
    end
  end

  def relation_tags
    @relation_tags ||= @package.relation_tags(self)
  end
end
