# coding: UTF-8

require 'rexml/document'

require 'zipruby'

require 'package_part'

class Package

  # 開いたファイルのパスです。
  attr_reader :file_path, :initialized_parts

  # 新しいインスタンスの初期化を行います。
  def initialize(file_path)
    @file_path = file_path
    @archive = Zip::Archive.open(file_path.encode('Shift_JIS'))
    @initialized_parts = []
    @parts = {}
  end

  def self.open(file_path)
    package = Package.new(file_path)
    if block_given?
      begin
        yield package
      ensure
        package.close
      end
    else
      package
    end
  end

  # ファイルの操作を終了し、ファイルを開放します。
  def close    
    @initialized_parts.select(&:changed?).each do |part|
      @archive.replace_buffer(part.part_uri, part.xml_document.to_s)
    end
    @archive.close
  end
  
  def root
    part('')
  end

  def part(uri)
    @parts[uri] ||= PackagePart.new(self, uri)
  end

  def xml_document(part)
    @archive.fopen(part.part_uri){|file| REXML::Document.new(file.read) }
  rescue REXML::ParseException
  end

  def relation_tags(part)
    @archive.fopen(part.rels_uri){|file| REXML::Document.new(file.read) }.get_elements('//Relationship')
  end
end