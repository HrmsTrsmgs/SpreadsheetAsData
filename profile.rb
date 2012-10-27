# coding:UTF-8
$: <<  './lib'
require './lib/work_book'
require 'benchmark'
#require 'profiler'

#Profiler__.start_profile
WorkBook.open('./spec/test_data/テーブル.xlsx') do |book|
  puts Benchmark.measure {
    book.Sheet3.A1
  }
end
#Profiler__.print_profile(File.open('prof.log', 'a'))