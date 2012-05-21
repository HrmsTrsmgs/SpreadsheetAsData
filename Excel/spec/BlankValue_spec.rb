# -*- encoding: Shift_JIS -*- 
require File.dirname(__FILE__) + '/../lib/WorkBook'
describe BlankValue do
	def test_file(file_name)
		return '.\spec\test_data\\' + file_name + '.xlsx'
	end
	
	after(:all) do
		book.close
	end
	
	subject {sheet.B1}
	let(:book) { WorkBook.open(test_file('Book1')) }
	let(:sheet) { book.Sheet1 } 
	
	
	describe '#==' do
		it 'で比較して""と同じということになる。' do
			subject.should == ''
		end
		it 'で比較して0と同じということになる。' do
			subject.should == 0
		end
		it 'で比較して0以外の整数とはと同じということにらない。' do
			subject.should_not == 1
			subject.should_not == -1
		end
		it 'で比較して0以外の少数とはと同じということにらない。' do
			subject.should_not == 0.01
			subject.should_not == -0.01
		end
		it 'で比較して''以外の少数とはと同じということにらない。' do
			subject.should_not == 'a'
		end
	end
	describe '#+' do
		it 'で0扱いで数字と足すことができる' do
			(subject + 10).should == 10
			(subject + -5).should == -5
		end
		it 'で0扱いで数字に足されることができる' do
			(10 + subject).should == 10
			(-5 + subject).should == -5
		end
		it 'で""扱いで文字列と足すことができる' do
			(subject + 'abc').should == 'abc'
		end
		it 'で""扱いで文字列に足されることができる' do
			('abc' + subject).should == 'abc'
		end
	end
	
	describe '#*' do
		it 'で0扱いで数字と掛けることができる' do
			(subject * 10).should == 0
		end
		it 'で0扱いで数字に掛けられることができる' do
			(10 * subject).should == 0
		end
		it 'で""扱いで文字列に足されることができる' do
			('abc' * subject).should == ''
		end
	end
	
	describe '数値演算' do
		it 'を行うことができる' do
			Math::sin(subject).should == 0
			Math::cos(subject).should == 1
		end
	end
end