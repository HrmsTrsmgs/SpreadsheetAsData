# -*- encoding: Shift_JIS -*- 

$: << File.dirname(__FILE__) + '/../lib'

require 'work_book'

def test_file(file_name)
  "./spec/test_data/#{file_name}.xlsx"
end