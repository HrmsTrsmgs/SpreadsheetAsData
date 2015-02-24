# coding: UTF-8
require 'rexml/document'

require 'work_sheet'
require 'package'
require 'shared_strings'

# Excelブックを扱うクラスです。
class WorkBook

  # 共有文字列をさすリレーションのIDです。
  SHARED_STRING_ID = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
  
  # ブックドキュメントをさすリレーションのIDです。
  DOCUMENT_ID = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'

  # ブック内の文字列を出力するエンコードです。
  attr_reader :encoding

  # 新しいインスタンスの初期化を行います。
  
  def initialize(file_path, encoding)
    @package = Package.new(file_path)
    @encoding = encoding
    @part = @package.root.relation(DOCUMENT_ID)
  end

  # Excelファイルのパスを指定し、開きます。
  # =[PARAM]file_path
  # =開くファイルのパスを指定します。
  def self.open(file_path, encoding = file_path.encoding)
    book = WorkBook.new(file_path, encoding)
    if block_given?
      begin
        yield book
      ensure
        book.close
      end
    else
      book
    end
  end

  # ファイルの操作を終了し、ファイルを開放します。
  def close
    @package.close
  end
  
  # 開いたExcelファイルのパスです。
  def file_path
    @package.file_path
  end
  
  # このブックで利用する共有文字列を取得します。
  def shared_strings
    @shared_strings ||= SharedStrings.new(@part.relation(SHARED_STRING_ID))
  end

  # ブックが持つシートを取得します。
  def sheets
    @sheets ||=
      @part.xml_document.root.get_elements('./sheets/sheet')
        .map{ |tag| WorkSheet.new(tag, @part.relation(tag).xml_document, self, @part.relation(tag))}
  end

  # 名前を指定し、ブックが持つシートを取得します。
  def [](name)
    sheets.find{|sheet| sheet.name == name.to_s.encode(encoding)}
  end

  # 指定したメソッド名が定義されていない場合に呼び出されます。
  def method_missing(method_name, *args)
    self[method_name] || super
  end
end