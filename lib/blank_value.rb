# coding: UTF-8

class BlankValue < Numeric
  def ==(other)
    other == 0 || other == ''
  end

  def inspect
    '{blank}'
  end

  def to_s
    ''
  end

  def method_missing(method_name, *args)
    begin
      0.__send__(method_name, *args)
    rescue
      ''.__send__(method_name, *args)
    end
  end
end