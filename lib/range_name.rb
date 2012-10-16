# coding: UTF-8

class RangeName

  def self.valid?(name = nil)
    name ? self.new(name).valid? : RangeNameValidator.new
  end  

  attr_reader :upper_left, :lower_right

  def initialize(*names)
    case names.join(':')
    when /^([A-Z]+)(\d+)(:|_)([A-Z]+)(\d+)$/
      @columns = false
      column_names = [$1, $4]
      row_nums = [$2.to_i, $5.to_i]
    when /^([A-Z]+)(:|_)([A-Z]+)$/
      @columns = true
      column_names = [$1, $3]
      row_nums = [1, 1048576]
    else
      column_names = [nil]
    end

    if column_names.all?
      @valid = true
      @upper_left = CellName.new(column_names.min + row_nums.min.to_s)
      @lower_right = CellName.new(column_names.max + row_nums.max.to_s)
      invalid! unless @upper_left.valid? && @lower_right.valid?
    end
  end
  
  def columns?
    @columns
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
    case
    when columns?
      "#{upper_left.column_name}:#{lower_right.column_name}"
    when valid?
     "#@upper_left:#@lower_right"
    else
      '{invalid range}'
    end
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