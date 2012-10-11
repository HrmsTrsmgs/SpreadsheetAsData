# coding: UTF-8

class CellName
  attr_reader :column_name, :row_num
  def initialize(name)
    if @valid = name =~ /^([A-Z]+)(\d+)$/
      @column_name = $1
      @row_num = $2.to_i
      
      if 2**14 < column_num || 2**20 < @row_num
        @valid = false
        @column_name = nil
        @row_num = nil
      end
    end
  end
  
  def valid?
    @valid
  end
  
  def column_num
    [*'A'..column_name].size
  end
  
  private
  
  
end