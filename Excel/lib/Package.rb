require 'fileutils'

require 'zip/zipfilesystem'
require 'rexml/document'

require 'package_part'

class Package

  # 開いたファイルのパスです。
  attr_reader :file_path

  # 新しいインスタンスの初期化を行います。
  def initialize(file_path)
    @file_path = file_path
    unzip_file
  end

  # ファイルの操作を終了し、ファイルを開放します。
  def close
    FileUtils.remove_entry(unziped_dir_path.encode("Shift_JIS"))
  end

  def part(uri)
    PackagePart.new(self, uri)
  end

  def part_path(uri)
    (unziped_dir_path + uri).gsub('//', '/')
  end

  private
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

  # 編集用に解凍されたフォルダのパスを取得します。
  def unziped_dir_path
    File.dirname(slashed_file_path) + "/tmp_" + File.basename(slashed_file_path) + "/"
  end

  # Ruby上でファイルパスとして認識される、スラッシュ区切りのファイルパスを取得します。
  def slashed_file_path
    file_path.gsub('\\', '/')
  end
end