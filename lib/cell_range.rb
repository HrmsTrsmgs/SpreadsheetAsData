# coding: UTF-8

class CellRange
  
  attr_reader :sheet
  
  def initialize(corner1, corner2, sheet)
    @corner1 = sheet.cell(corner1)
    @corner2 = sheet.cell(corner2)
    @sheet = sheet
  end
  
  def book
    sheet.book
  end
  
  def upper_left_cell
    @corner1
  end
  
  def lower_right_cell
    @corner2
  end
  
  def all
    /^[A-Z]+(\d+):[A-Z]+(\d+)$/ =~ to_s
    Array.new($2.to_i - $1.to_i)
  end
  
  def to_s
    "#@corner1:#@corner2"
  end
  
  alias ref to_s
end