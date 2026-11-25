# coding: UTF-8

$: << File.dirname(__FILE__) + '/../lib'
Encoding.default_external = Encoding::UTF_8

require 'work_book'

RSpec.configure do |config|
  config.expect_with :rspec do |expectations|
    expectations.syntax = [:should, :expect]
  end
end

TEST_FILE_DIRECTORY = '../TestData/'

def test_file(basename)
  directory = basename == 'テーブル' ? '空白数値なし/' : ''
  "#{TEST_FILE_DIRECTORY}#{directory}#{basename}.xlsx"
end

class TestFile
  @opened_books = []
  @copy_paths = {}

  class << self
    basenames = Dir.glob("#{TEST_FILE_DIRECTORY}*.xlsx").map{|file| File.basename(file, '.xlsx') }
    basenames.each do |basename|
      basename_downcase = basename.downcase.gsub('-', '_')
      define_method basename_downcase do
        open_book(send("#{basename_downcase}_path"))
      end

      define_method "#{basename_downcase}_copy" do
        open_book(send("#{basename_downcase}_copy_path"))
      end

      define_method("#{basename_downcase}_path") do
        test_file(basename)
      end

      define_method "#{basename_downcase}_copy_path" do
        @copy_paths[basename] ||= begin
          path = test_file("#{basename}_Copy")
          FileUtils.cp(send("#{basename_downcase}_path"), path)
          path
        end
      end
    end

    def open_book(path, *options)
      book = WorkBook.open(path, *options)
      @opened_books << book
      book
    end

    def close
      @opened_books.each(&:close)
      @opened_books.clear
      FileUtils.rm_f(@copy_paths.values)
      @copy_paths.clear
    end
  end
end
