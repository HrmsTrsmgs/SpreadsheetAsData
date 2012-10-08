# coding: UTF-8

class DataRow
  def initialize(cell_range, row_num, left_num)
    @cell_range = cell_range
    @row_num = row_num
  end
  
  def cell_value(table_column_name)
    cell_column_name = 'A'
    column_index = @cell_range.column_names.index{|name| name.to_s == table_column_name.to_s}
    (@cell_range.upper_left_cell.column_num + column_index - 1).times { cell_column_name.succ! }
    cell_name = cell_column_name +  @row_num.to_s
    @cell_range.sheet.cell_value(cell_name)
  end
  
  def method_missing(method_name)
    cell_value(method_name) || super
  end
end