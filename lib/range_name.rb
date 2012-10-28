# coding: UTF-8

class RangeName

  def self.valid?(name = nil)
    name ? self.new(name).valid? : RangeNameValidator.new
  end  

  attr_reader :sheet, :upper_left, :lower_right

  def initialize(*names)
    
    match = names.join(':') =~ /^(?:(.*?)!)?([A-Z]+)(\d+)?(?::|_)([A-Z]+)(\d+)?$/
    if match && $3 && $5
      @valid = true
      @columns = false
      row_nums = [$3.to_i, $5.to_i]
    elsif match && !$3 && !$5
      @valid = true
      @columns = true
      row_nums = [1, 1048576]
    else
      @valid == false
    end
    @sheet = $1
    column_names = [$2, $4]
    
    if @valid
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