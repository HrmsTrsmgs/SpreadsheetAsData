# coding: UTF-8

require 'forwardable'

require_relative 'data_row'
require_relative 'range_name'


class CellRange
  extend Forwardable

  attr_reader :sheet, :upper_left_cell, :lower_right_cell

  def_delegator(:sheet, :book)

  def initialize(corner1, corner2, sheet)
    @upper_left_cell = sheet.cell(corner1)
    @lower_right_cell = sheet.cell(corner2)
    @sheet = sheet
  end

  def all
    data_row_numbers.select { |row_num| row_has_data?(row_num) }.
      map { |row_num| DataRow.new(self, row_num) }
  end

  def where(exp)
    result = all
    exp.each do |key, value|
      result =
        result.select do |row|
          case value
          when Array
            value.include? row.cell_value(key.to_s)
          else
            value === row.cell_value(key.to_s)
          end
        end
    end
    result
  end

  def order(sort)
    key, order = sort.split
    result = all.sort_by{|row| row.cell_value(key) }

    if order && order.downcase == 'desc'
      result.reverse
    else
      result
    end
  end

  def column_names
    column_name_range.map do |column|
      sheet.cell_value(column + upper_left_cell.row_num.to_s).to_sym
    end
  end

  def to_s
    "#@upper_left_cell:#@lower_right_cell"
  end

  alias ref to_s

private
  def name
    RangeName.new(to_s)
  end

  def data_row_numbers
    bounds = (upper_left_cell.row_num + 1)..lower_right_cell.row_num
    sheet.real_row_nums.select { |row_num| bounds.cover?(row_num) }.uniq.sort
  end

  def row_has_data?(row_num)
    column_name_range.any? { |column_name| !sheet.blank?(column_name + row_num.to_s) }
  end

  def column_name_range
    upper_left_cell.column_name..lower_right_cell.column_name
  end
end
