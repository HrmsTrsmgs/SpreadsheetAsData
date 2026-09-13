# coding:UTF-8
$: <<  './lib'
require './lib/work_book'
require 'benchmark'
#require 'profiler'

#Profiler__.start_profile
WorkBook.open(File.expand_path('../TestData/空白数値なし/テーブル.xlsx', __dir__)) do |book|
  puts Benchmark.measure {
    book.Sheet3.A1
  }
end
#Profiler__.print_profile(File.open('prof.log', 'a'))
