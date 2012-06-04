# coding: UTF-8
require 'rexml/document'

require 'work_sheet'
require 'package'
require 'shared_strings'

# Excelブックを扱うクラスです。
class WorkBook

  SHARED_STRING_ID = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'

  attr_reader :encoding, :shared_strings

  # 開いたExcelファイルのパスです。
  def file_path
    @package.file_path
  end

  # 新しいインスタンスの初期化を行います。
  def initialize(file_path, encoding)
    @package = Package.new(file_path)
    @encoding = encoding
    @part = @package.part('xl/workbook.xml')
    @shared_strings = SharedStrings.new(@part.relation(SHARED_STRING_ID))
  end

  # Excelファイルのパスを指定し、開きます。<br/>
  # <br/>
  # ==== <b>[PARAM]file_path</b><br/>
  #  開くファイルのパスを指定します。
  def self.open(file_path, encoding = file_path.encoding)
      book = WorkBook.new(file_path, encoding)
      if block_given?
        yield book
        book.close
      else
        book
      end
  end

  # ファイルの操作を終了し、ファイルを開放します。
  def close
    @package.close
  end

  # ブックが持つシートを取得します。
  def sheets
    @sheets ||=
      @part.xml_document.elements.to_a('//sheet')
        .map{ |tag| WorkSheet.new(tag, @part.relation(tag).xml_document, self, @part.relation(tag))}
  end

  # 名前を指定し、ブックが持つシートを取得します。
  def [](name)
    sheets.find{|sheet| sheet.name == name.to_s.encode(encoding)}
  end

  # 指定したメソッド名が定義されていない場合に呼び出されます。
  def method_missing(method_name)
    self[method_name] || super
  end
end