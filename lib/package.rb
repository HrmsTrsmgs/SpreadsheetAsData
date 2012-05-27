require 'zipruby'
require 'rexml/document'

require 'package_part'

class Package

  # 開いたファイルのパスです。
  attr_reader :file_path

  # 新しいインスタンスの初期化を行います。
  def initialize(file_path)
    @file_path = file_path
    @archive = Zip::Archive.open(file_path)
  end

  # ファイルの操作を終了し、ファイルを開放します。
  def close
    @archive.close()
  end

  def part(uri)
    PackagePart.new(self, uri)
  end
  
  # ブック情報を記述してあるWorkBook.xmlドキュメントを取得します。
  def xml_document(part_uri)
    #workbook.xmlのパスは変更するとExcelでも起動できなくなるため、変更には対応しません。
    part_uri.gsub!(/^\//, '')
    @archive.fopen(part_uri){|file| REXML::Document.new(file.read) }
  end

  def relation_tags(part_uri)
    part_uri.gsub!(/^\//, '')
    @archive.fopen(part_rels_file_path(part_uri)){|file| REXML::Document.new(file.read) }.elements.to_a('//Relationship')
  end

  private
  def part_rels_file_path(part_uri)
    File.dirname(part_uri) + '/_rels/' + File.basename(part_uri) + '.rels'
  end
end