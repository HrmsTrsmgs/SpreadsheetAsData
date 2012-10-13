# coding: UTF-8

class RangeName

  def self.valid?(name)
    self.new(name).valid?
  end

  attr_reader :start_cell, :end_cell

  def initialize(name)
    if name =~ /^([A-Z]+\d+):([A-Z]+\d+)$/
      @valid = true
      @start_cell = CellName.new($1)
      @end_cell = CellName.new($2)
      
      invalid! unless @start_cell.valid? && @end_cell.valid?
    end
  end
  
  def valid?
    @valid
  end
  
  private
    def invalid!
      @valid = false
      @start_cell = nil
      @end_cell = nil
    end
end