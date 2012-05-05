require 'zip/zipfilesystem'
require 'fileutils'
require 'rexml/document'

class PackagePart
	def initialize(package, part_uri)
		@package = package
		@part_uri = part_uri
	end
	
	def file_path
		@package.part_path(@part_uri)
	end
	
	def rels_file_path
		File.dirname(file_path) + '/_rels/' + File.basename(file_path) + '.rels'
	end
	
	# ブック情報を記述してあるWorkBook.xmlドキュメントを取得します。
	def xml_document
		#workbook.xmlのパスは変更するとExcelでも起動できなくなるため、変更には対応しません。
		@xml_document ||= File.open(file_path) {|file| REXML::Document.new(file) }
	end
	
	def relation_tags
		File.open(rels_file_path) {|file| REXML::Document.new(file) }.elements.to_a('//Relationship')
	end
	
	def relation_cache
		if !@hash then
			relations_array =
				relation_tags
					.map{|tag| 
						[tag.attributes['Id'],
						PackagePart.new(
							@package,
							File.dirname(@part_uri) +"/" + tag.attributes['Target'])
						]
					}
			@hash = Hash[*relations_array.flatten(1)]
		end
		
		@hash
	end
	
	def relation(relation_id_or_tag)
			if relation_id_or_tag.respond_to?(:attributes) && relation_id_or_tag.attributes['r:id'] then
				relation_cache[relation_id_or_tag.attributes['r:id']]
			else
				relation_cache[relation_id_or_tag]
			end
	end
	
	def relations
		if !@relations then
			relations_array =
				relation_tags
					.map{|tag| 
						[tag.attributes['Id'],
						PackagePart.new(
							@package,
							File.dirname(@part_uri) +"/" + tag.attributes['Target'])
						]
					}
			@relations = PackageRelations.new(Hash[*relations_array.flatten(1)])
		end
		
		@relations
	end
end