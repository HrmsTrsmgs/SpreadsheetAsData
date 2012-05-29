# -*- encoding: UTF-8 -*- 
require 'pathname'
require 'rexml/document'

require 'zipruby'

require 'package_part'

class Package

  # �J�����t�@�C���̃p�X�ł��B
  def file_path
    @file_path.to_s
  end

  # �V�����C���X�^���X�̏��������s���܂��B
  def initialize(file_path)
    @file_path = Pathname(file_path)
    @archive = Zip::Archive.open(file_path)
  end

  # �t�@�C���̑�����I�����A�t�@�C�����J�����܂��B
  def close
    @archive.close()
  end

  def part(uri)
    PackagePart.new(self, uri)
  end
  
  # �u�b�N�����L�q���Ă���WorkBook.xml�h�L�������g���擾���܂��B
  def xml_document(part_uri)
    #workbook.xml�̃p�X�͕ύX�����Excel�ł��N���ł��Ȃ��Ȃ邽�߁A�ύX�ɂ͑Ή����܂���B
    @archive.fopen(part_uri.to_s){|file| REXML::Document.new(file.read) }
  end

  def relation_tags(part_uri)
    @archive.fopen(part_rels_file_path(part_uri).to_s){|file| REXML::Document.new(file.read) }.elements.to_a('//Relationship')
  end

  private
  def part_rels_file_path(part_uri)
    part_uri.dirname + '_rels/' + (part_uri.basename.to_s + '.rels')
  end
end