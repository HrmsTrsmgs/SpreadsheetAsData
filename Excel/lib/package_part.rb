require 'zip/zipfilesystem'
require 'rexml/document'

class PackagePart
  def initialize(package, part_uri)
    @package = package
    @part_uri = part_uri
  end

  def file_path
    @package.part_path(@part_uri)
  end

  def rels_file_path
    File.dirname(file_path) + '/_rels/' + File.basename(file_path) + '.rels'
  end

  # ブック情報を記述してあるWorkBook.xmlドキュメントを取得します。
  def xml_document
    #workbook.xmlのパスは変更するとExcelでも起動できなくなるため、変更には対応しません。
    @xml_document ||= File.open(file_path) {|file| REXML::Document.new(file) }
  end

  def relation_tags
    @relation_tags ||= File.open(rels_file_path) {|file| REXML::Document.new(file) }.elements.to_a('//Relationship')
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
end