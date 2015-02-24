require 'fileutils'

require 'zip/zipfilesystem'
require 'rexml/document'

require 'package_part'

class Package

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
  
  # 開いたファイルのパスです。
  attr_reader :file_path, :initialized_parts

  # 新しいインスタンスの初期化を行います。
  def initialize(file_path)
    @file_path = file_path
    @initialized_parts = []
    @parts = {}
    unzip_file
    @file = File.open(slashed_file_path)
  end

  # ファイルの操作を終了し、ファイルを開放します。
  def close
    return if @closed
    @initialized_parts.select(&:changed?).each do |part|
      File.write(unziped_dir_path + part.part_uri, part.xml_document.to_s)
    end
    FileUtils.remove_entry(unziped_dir_path.encode("Shift_JIS")) if Dir.exist?(unziped_dir_path)
    @file.close
    @closed = true
  end
  
  def root
    part('')
  end

  def part(uri)
    @parts[uri] ||= PackagePart.new(self, uri)
  end
  
  # ブック情報を記述してあるWorkBook.xmlドキュメントを取得します。
  def xml_document(part)
    #workbook.xmlのパスは変更するとExcelでも起動できなくなるため、変更には対応しません。
    archive_as_document(part.part_uri)
    rescue REXML::ParseException
  end

  def relation_tags(part)
    archive_as_document(part.rels_uri).get_elements('/Relationships/Relationship')
  end
  
  private
  def archive_as_document(path)
    File.open(part_file_path(path)) {|file| REXML::Document.new(file) }
  end
  # ファイルのzip圧縮を解凍し、編集可能とします。
  def unzip_file
    Zip::ZipInputStream.open(slashed_file_path) do |stream|
      while ziped_file = stream.get_next_entry()
        dir_name = File.dirname(ziped_file.name)
        FileUtils.makedirs(unziped_dir_path + dir_name)
        ziped_file_name =  unziped_dir_path + ziped_file.name
        unless ziped_file_name.match(/\/$/)
          File.open(ziped_file_name, "w+b") do |written|
            written.puts(stream.read())
          end
        end
      end
    end
  end
  
  def zip_file
    Zip::ZipFile.open(slashed_file_path, Zip::ZipFile::CREATE) do |zip_file|
    	src_paths.each do |src_path|
      	src_path = File.expand_path(src_path)
      zip_file.add(File.basename(src_path), src_path)
    end
  end
  end

  # 編集用に解凍されたフォルダのパスを取得します。
  def unziped_dir_path
    File.dirname(slashed_file_path) + "/tmp_" + File.basename(slashed_file_path) + "/"
  end

  # Ruby上でファイルパスとして認識される、スラッシュ区切りのファイルパスを取得します。
  def slashed_file_path
    file_path.gsub('\\', '/')
  end

  def part_path(uri)
    (unziped_dir_path + uri).gsub('//', '/')
  end
  def part_file_path(part_uri)
    part_path(part_uri)
  end

  def part_rels_file_path(part_uri)
    File.dirname(part_file_path(part_uri)) + '/_rels/' + File.basename(part_file_path(part_uri)) + '.rels'
  end
end