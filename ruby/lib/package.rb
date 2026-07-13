require 'fileutils'

require 'zip'
require 'rexml/document'

require_relative 'package_part'

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
    changed_parts = @initialized_parts.select(&:changed?)
    changed_parts.each do |part|
      File.write(unziped_dir_path + part.part_uri, part.xml_document.to_s)
    end
    @file.close
    if changed_parts.any?
      File.delete(slashed_file_path)
      zip_file
    end
    FileUtils.remove_entry(unziped_dir_path) if Dir.exist?(unziped_dir_path)
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
    Zip::File.open(slashed_file_path) do |zip|
      zip.each do |file|
        entry_name = file.name.force_encoding(Encoding::UTF_8)
        dir_name = File.dirname(entry_name)
        FileUtils.makedirs(unziped_dir_path + dir_name)
        ziped_file_name =  unziped_dir_path + entry_name
        unless ziped_file_name.match(/\/$/)
          File.open(ziped_file_name, "w+b") do |written|
            stream = file.get_input_stream
            written.write(stream.read)
            stream.close
          end
        end
      end
    end
  end
  
  def zip_file
    
    Zip::File.open(slashed_file_path, Zip::File::CREATE) do |zip_file|
    	Dir::glob(unziped_dir_path + "**/*", File::FNM_DOTMATCH).each do |src_path|
        next if ['.', '..'].include?(File.basename(src_path))

        zip_path = src_path.delete_prefix(unziped_dir_path)
        if File.file?(src_path)
          zip_file.add(zip_path, src_path)
        else
          zip_file.mkdir(zip_path)
        end
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
