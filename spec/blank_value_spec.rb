# coding: UTF-8

require 'spec_helper'

describe BlankValue do
  subject { sheet.B1 }
  let(:book) { TestFile.book1 }
  let(:sheet) { book.Sheet1 } 

  after do
    TestFile.close
  end

  describe '#==' do
    it 'で比較すると""と同じということになる。' do
      subject.should == ''
    end

    it 'で比較すると0と同じということになる。' do
      subject.should == 0
    end

    it 'で比較すると0以外の整数とはと同じということにらない。' do
      subject.should_not == 1
      subject.should_not == -1
    end

    it 'で比較すると0以外の小数とはと同じということにらない。' do
      subject.should_not == 0.01
      subject.should_not == -0.01
    end

    it 'で比較すると''以外の文字列とはと同じということにらない。' do
      subject.should_not == 'a'
      subject.should_not == 'あいう'
    end
  end

  describe '#+' do
    it 'で0扱いで数字と足すことができる。' do
      (subject + 10).should == 10
      (subject + -5).should == -5
    end

    it 'で0扱いで数字に足されることができる。' do
      (10 + subject).should == 10
      (-5 + subject).should == -5
    end

    it 'で""扱いで文字列と足すことができる。' do
      (subject + 'abc').should == 'abc'
    end

    it 'で""扱いで文字列に足されることができる。' do
      ('abc' + subject).should == 'abc'
    end
  end

  describe '#*' do
    it 'で0扱いで数字と掛けることができる。' do
      (subject * 10).should == 0
    end

    it 'で0扱いで数字に掛けられることができる。' do
      (10 * subject).should == 0
    end

    it 'で""扱いで文字列に足されることができる。' do
      ('abc' * subject).should == ''
    end
  end

  describe '数値演算' do
    it 'に利用することができる。' do
      Math::sin(subject).should == 0
      Math::cos(subject).should == 1
    end
  end

  describe '#inspect' do
    it 'は{blank}を返す' do
      subject.inspect.should == '{blank}'
    end
  end

  describe '#to_s' do
    it 'は""を返す' do
      subject.to_s.should == ''
    end
  end
end