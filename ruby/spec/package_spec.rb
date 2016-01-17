# Coding: utf-8

require_relative 'spec_helper'
require 'package'

describe Package do
  subject { Package.open(TestFile.book1_path) }
  
  after do
    subject.close
    TestFile.close
  end
  
  describe '::open' do
    it 'はブロック引数をとります。' do
      Package.open(TestFile.book1_path) do |package|
        package.part('xl/workbook.xml').xml_document.should_not be_nil
      end
    end
    
    it 'はファイルを束縛します。' do
      package = Package.open(TestFile.book1_copy_path)
      expect { File.delete(TestFile.book1_copy_path) }.to raise_error Errno::EACCES
      package.close
    end
    
    it 'はブロック内でファイルを束縛します。' do
      Package.open(TestFile.book1_copy_path) do |package|
        expect { File.delete(TestFile.book1_copy_path) }.to raise_error Errno::EACCES
      end
    end
    
    it 'はブロック終了時にファイルを開放します。' do
      Package.open(TestFile.book1_copy_path) do |package|
      end
      expect { File.delete(TestFile.book1_copy_path) }.to_not raise_error
    end
    
    it 'は例外発生時もブロック終了時にファイルを開放します。' do
      begin
        Package.open(TestFile.book1_copy_path) do |package|
          raise Exception
        end
      rescue Exception
      end
      expect { File.delete(TestFile.book1_copy_path) }.to_not raise_error
    end
  end
  
  describe '#close' do
    it 'はファイルを開放します。' do
      subject.close
      expect { File.delete(TestFile.book1_copy_path) }.to_not raise_error
    end
  end
  
  describe '#root' do
    it 'はトップレベルのファイルを取得します。' do
      subject.root.part_uri.should == ''
      subject.root.relation('rId1').part_uri.should == ('xl/workbook.xml')
    end
  end
  
  describe '#part' do
    it 'によりPackagePartが取得できます。' do
      subject.part('xl/workbook.xml').part_uri.should == 'xl/workbook.xml'
    end
    
    it 'により取得されたPartは毎回同一になります。' do
      subject.part('xl/workbook.xml').should equal subject.part('xl/workbook.xml')
    end
  end
  
  describe '#xml_document' do
    it 'によりPackagePartから対応するXMLドキュメントがが取得できます。' do
      subject.xml_document(subject.part('xl/workbook.xml')).root.name.should == 'workbook'
    end
  end
  
  describe '#relation_tags' do
    it 'によりPackagePartから対応するリレーションが取得できます。' do
      subject.relation_tags(subject.part('xl/workbook.xml')).map{|tag| tag.attributes['Id']}.sort.should == ['rId1', 'rId2', 'rId3', 'rId4', 'rId5', 'rId6']
    end
  end
end