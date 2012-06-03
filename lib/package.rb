# coding: UTF-8

require 'zipruby'
require 'rexml/document'

require 'package_part'

class Package

  # 開いたファイルのパスです。
  attr_reader :file_path, :initialized_parts

  # 新しいインスタンスの初期化を行います。
  def initialize(file_path)
    @file_path = file_path
    @archive = Zip::Archive.open(file_path.encode('Shift_JIS'))
    @initialized_parts = []
  end

  # ファイルの操作を終了し、ファイルを開放します。
  def close
    @initialized_parts.select(&:changed?).each do |part|
      @archive.replace_buffer(part.part_uri, part.xml_document.to_s)
    end
    @archive.close
  end

  def part(uri)
    PackagePart.new(self, uri)
  end
  
  # ブック情報を記述してあるWorkBook.xmlドキュメントを誌・ｵます。
  def xml_document(part)
    #workbook.xmlのパスは変更するとExcelでも起動できなくなるため、変更には対応しません。
    @archive.fopen(part.part_uri){|file| REXML::Document.new(file.read) }
  end

  def relation_tags(part)
    @archive.fopen(part_rels_file_path(part.part_uri)){|file| REXML::Document.new(file.read) }.elements.to_a('//Relationship')
  end

  private
  def part_rels_file_path(part_uri)
    File.dirname(part_uri) + '/_rels/' + File.basename(part_uri) + '.rels'
  end
end