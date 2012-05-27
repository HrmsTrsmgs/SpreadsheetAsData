require 'pathname'
require 'rexml/document'

require 'zipruby'

require 'package_part'

class Package

  # 開いたファイルのパスです。
  def file_path
    @file_path.to_s
  end

  # 新しいインスタンスの初期化を行います。
  def initialize(file_path)
    @file_path = Pathname(file_path)
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
    @archive.fopen(part_uri.to_s){|file| REXML::Document.new(file.read) }
  end

  def relation_tags(part_uri)
    @archive.fopen(part_rels_file_path(part_uri).to_s){|file| REXML::Document.new(file.read) }.elements.to_a('//Relationship')
  end

  private
  def part_rels_file_path(part_uri)
    part_uri.dirname + '_rels/' + (part_uri.basename.to_s + '.rels')
  end
end