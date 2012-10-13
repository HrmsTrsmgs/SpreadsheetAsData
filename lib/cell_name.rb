# coding: UTF-8

class CellName

  def self.valid?(name = nil)
     name ? self.new(name).valid? : CellNameValidator.new
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
    [*'A'..column_name].size if valid?
  end
  
  def to_s
    @column_name + @row_num.to_s
  end
  
  def eql?(other)
    column_name == other.column_name && row_num == other.row_num
  end
  
  def hash
    [column_name, row_num].hash
  end
  
private
  def invalid!
    @valid = false
    @column_name = nil
    @row_num = nil
  end
  
  class CellNameValidator
    def ===(other)
      CellName.valid?(other)
    end
  end
end