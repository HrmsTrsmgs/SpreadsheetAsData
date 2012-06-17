# coding: UTF-8

require 'spec_helper'
require 'package'

describe PackagePart do
  let(:package){ Package.open(TestFile.book1_path) }
  subject { package.part('xl/workbook.xml') }
  let(:printer_settings1){ package.part('xl/printerSettings/printerSettings1.bin')}
  let(:written){ Package.open(TestFile.book1_copy_path) }
  let(:written_part){ written.part('xl/workbook.xml') }
  
  after do
    package.close
    written.close
    TestFile.close
  end
  
  describe '#part_uri' do
    it 'で指定したパートURIが取得できます。' do
      subject.part_uri.should == 'xl/workbook.xml'
    end
  end
  
  describe '#rels_uri' do
    it 'で指定したパートのリレーションファイルURIが取得できます。' do
      subject.rels_uri.should == 'xl/_rels/workbook.xml.rels'
    end
  end
  
  describe '#changed?' do
    it 'の初期値はfalseです。' do
      subject.should_not be_changed
    end
    
    it 'は変更された場合にtrueとなります。' do
      written_part.change
      written_part.should be_changed
    end
  end
  
  describe '#changed' do
    it 'は#changed?を変更します。' do
      written_part.change
      written_part.should be_changed
    end
  end

  describe '#xml_document' do
    it 'はXMLドキュメントを取得します。' do
      subject.xml_document.root.name.should == 'workbook'
    end
    
    it 'はXMLではないパートの場合はnilを返します。' do
      printer_settings1.xml_document.should be_nil
    end
  end
  
  describe '#relation' do
    it 'はXML名前空間を受け付け、該当する結果を一つ返します。' do
      subject.relation('http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme').part_uri.should == 'xl/theme/theme1.xml'
    end
    
    it 'はRelation IDを受け付け、該当する結果を一つ返します。' do
      subject.relation('rId2').part_uri.should == 'xl/worksheets/sheet2.xml'
    end
    
    it 'はRelation IDを含むタグを受け付け、該当する結果を一つ返します。' do
      tag = REXML::Element.new('sheet')
      tag.attributes['r:id'] = 'rId3'
      subject.relation(tag).part_uri.should == 'xl/worksheets/sheet3.xml'
    end
  end
end