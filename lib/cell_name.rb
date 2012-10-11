# coding: UTF-8

class CellName
  attr_reader :is_cell_name, :column_name, :row_num
  def initialize(name)
    if @is_cell_name = name =~ /^([A-Z]+)(\d+)$/
      @column_name = $1
      @row_num = $2.to_i
      
      if 2**14 < column_num || 2**20 < @row_num
        @is_cell_name = false
        @column_name = nil
        @row_num = nil
      end
    end
  end
  
  def column_num
    [*'A'..column_name].size
  end
end