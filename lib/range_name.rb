# coding: UTF-8

class RangeName

  def self.valid?(name = nil)
    name ? self.new(name).valid? : RangeNameValidator.new
  end  

  attr_reader :upper_left, :lower_right

  def initialize(*names)
    case names.size
    when 1
      names[0] =~ /^([A-Z]+)(\d+)(:|_)([A-Z]+)(\d+)$/
      column_names = [$1, $4]
      row_nums = [$2.to_i, $5.to_i]
    when 2
      column_names = names.map{|name| CellName.new(name).column_name }
      row_nums = names.map{|name| CellName.new(name).row_num }
    end
    if column_names.all?
      @valid = true
      @upper_left = CellName.new(column_names.min + row_nums.min.to_s)
      @lower_right = CellName.new(column_names.max + row_nums.max.to_s)
      invalid! unless @upper_left.valid? && @lower_right.valid?
    end
  end

  def valid?
    @valid
  end
  
  def eql?(other)
    upper_left.eql?(other.upper_left) && lower_right.eql?(other.lower_right)
  end
  
  def hash
    [upper_left, lower_right].hash
  end
  
  def to_s
    valid? ? "#@upper_left:#@lower_right" : '{invalid range}'
  end
private
  def invalid!
    @valid = false
    @start_cell = nil
    @end_cell = nil
  end
  
  class RangeNameValidator
    def ===(other)
      RangeName.valid?(other)
    end
  end
end