# coding: UTF-8

class CellName

  def self.valid?(name)
    self.new(name).valid?
  end

  attr_reader :column_name, :row_num

  def initialize(name)
    if name =~ /^([A-Z]+)(\d+)$/
      @valid = true
      @column_name = $1
      @row_num = $2.to_i
      
      invalid! if 2**14 < column_num || 2**20 < @row_num
    end
  end
  
  def valid?
    @valid
  end
  
  def column_num
    [*'A'..column_name].size
  end
  
  private
    def invalid!
      @valid = false
      @column_name = nil
      @row_num = nil
    end
end