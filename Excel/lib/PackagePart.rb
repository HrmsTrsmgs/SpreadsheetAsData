require 'zip/zipfilesystem'
require 'fileutils'
require 'rexml/document'

class PackagePart
	def initialize(package, url)
		@package = package
		@url = url
		@file_path = (@package.unziped_dir_path + @url).gsub('//', '/')
		@rels_file_path = File.dirname(@file_path) + '/_rels/' + File.basename(@file_path) + '.rels'
	end
	
	# ブック情報を記述してあるWorkBook.xmlドキュメントを取得します。
	def xml_document
		if !@xml_document then
			#workbook.xmlのパスは変更するとExcelでも起動できなくなるため、変更には対応しません。
			File.open(@file_path) do |file|
				@xml_document = REXML::Document.new(file)
			end
		end
		return @xml_document
	end
	
	def relations
		if !@relations then
			rels_xml = nil;
			File.open(@rels_file_path) do |file|
				rels_xml = REXML::Document.new(file)
			end
			relations_array =
				rels_xml.elements.to_a('//Relationship')
					.map{|tag| 
						[tag.attributes['Id'],
						PackagePart.new(
							@package,
							File.dirname(@url) +"/" + tag.attributes['Target'])
						]
					}
			@relations = PackageRelations.new(Hash[*relations_array.flatten(1)])
		end
		
		return @relations
	end
	
	class PackageRelations
		def initialize(relations_hash)
			@hash = relations_hash
		end
		
		def [](relation_id_or_tag)
			if relation_id_or_tag.respond_to?(:attributes) && !relation_id_or_tag.attributes['r:id'].nil? then
				return @hash[relation_id_or_tag.attributes['r:id']]
			else
				return @hash[relation_id_or_tag]
			end
		end
	end
end