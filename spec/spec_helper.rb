# coding: UTF-8

$: << File.dirname(__FILE__) + '/../lib'

require 'work_book'

TEST_FILE_DIRECTORY = './spec/test_data/'

def test_file(basename)
  "#{TEST_FILE_DIRECTORY}#{basename}.xlsx"
end

class TestFile
  class << self
    basenames = Dir.glob("#{TEST_FILE_DIRECTORY}*.xlsx").map{|file| File.basename(file, '.xlsx') }
    
    basenames.each do |basename|
      p basename_downcase = basename.downcase.gsub('-', '_')
      define_method basename_downcase do
        book = WorkBook.open(send("#{basename_downcase}_path"))
        name = "@#{basename_downcase}"
        instance_variable_set(name, book) until instance_variable_get(name)
        book
      end
      
      define_method "#{basename_downcase}_copy" do
        book = WorkBook.open(send("#{basename_downcase}_copy_path"))
        name = "@#{basename_downcase}_copy"
        instance_variable_set(name, book) until instance_variable_get(name)
        book
      end
      
      define_method("#{basename_downcase}_path") do
        test_file(basename)
      end
      
      define_method "#{basename_downcase}_copy_path" do
        name = "@#{basename_downcase}_copy_path"      
        if not instance_variable_get(name)
          copy_path = test_file("#{basename}_Copy")
          instance_variable_set(name, copy_path)
          FileUtils.cp(send("#{basename_downcase}_path"), copy_path)
        end
        instance_variable_get(name)
      end
    end
    define_method :close do
      basenames.each do |basename|
        basename_downcase = basename.downcase.gsub('-', '_')
        book_name = "@#{basename_downcase}"
        if instance_variable_get(book_name)
          instance_variable_get(book_name).close
          instance_variable_set(book_name, nil)
        end
        book_name_copy = "@#{basename_downcase}_copy"
        if instance_variable_get(book_name_copy)
          instance_variable_get(book_name_copy).close
          instance_variable_set(book_name_copy, nil)
          
          FileUtils.rm(p send("#{basename_downcase}_copy_path"))
          instance_variable_set("@#{basename_downcase}_copy_path", nil)
        end
      end
    end
  end
end