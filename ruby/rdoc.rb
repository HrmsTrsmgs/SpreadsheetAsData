# coding: UTF-8
# 適当に作ったバッチです。
# TODO rdocと同じ引数をとるべきではないですかね。
require 'pathname'

# TODO tmpフォルダのパスがドキュメントにしっかりと表示されてしまってます。
# TODO コードをフォルダで整理したらどうするんでしょうね。
lib_dir = Pathname('lib')
tmp_dir = Pathname('lib_tmp')

tmp_dir.mkdir until tmp_dir.exist?

Pathname.glob(lib_dir + '*.rb') do |file|
  tmp = tmp_dir + file.relative_path_from(lib_dir)
  tmp.open('w') do |f|
    f.write(file.read(:encoding => 'UTF-8:Shift_JIS').gsub('UTF-8', 'Shift_JIS')) #TODO 置換処理は適当なのでいずれ直す。
  end
end

`rdoc lib_tmp -c Shift_JIS`.each_line {|line| puts line}

tmp_dir.rmtree
