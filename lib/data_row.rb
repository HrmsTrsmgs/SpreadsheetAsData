# coding: UTF-8

class DataRow
  def initialize(cell_range, row_num)
    @cell_range = cell_range
    @row_num = row_num
  end
  
  def cell_value(table_column_name)
    @cell_range.sheet.cell_value(column_name(table_column_name) +  @row_num.to_s)
  end
  
  private
  
  def column_name(table_column_name)
    column_num_to_name(@cell_range.upper_left_cell.column_num + column_index(table_column_name))
  end
  
  def column_index(table_column_name)
    @cell_range.column_names.index{|name| name.to_s == table_column_name.to_s}
  end
  
  
  def column_num_to_name(i)
    cell_column_name = 'A'
    (i - 1).times { cell_column_name.succ! }
    cell_column_name
  end
  
  def method_missing(method_name)
    cell_value(method_name) || super
  end
end