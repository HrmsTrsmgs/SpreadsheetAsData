# coding: UTF-8

$: << File.dirname(__FILE__) + '/../lib'

require 'work_book'

def test_file(file_name)
  "./spec/test_data/#{file_name}.xlsx"
end

def message_in_sjis
  class << self
    alias __it__ it
    
    def it(message)
      __it__(message.to_s.encode('Shift_JIS')){ yield if block_given? }
    end
  end
end