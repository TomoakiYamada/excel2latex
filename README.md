# Excelで作成した表をLaTeXコードに変換するJuliaコード
- version 0.1

## 使い方
- メインファイルはsrcにある`excel2latex.jl`です。
  - 2行目にあるpathを自分が作業している場所に設定してください。
- excel内にあるsample.xlsxを参考にして、作成したい表をExcelファイルにコピペしてください。
- `excel2latex.jl`を実行すると、outputフォルダに`table.tex`というファイルが作成されます。

## 事前準備
- Juliaを使っています。
  - 事前に[Julia](https://julialang.org/)をダウンロードしてインストールをしてください。
- XLSXというパッケージを使います。以下のコマンドでーパッケージを追加してください。
  - `using Pkg`
  - `Pkg.add("XLSX")`

## きれいな表を作成するためのTips
- 有効桁数などの見栄えはExcel側で整えたほうが楽です。
- 各セルのすべての数字が出力されるので、Excel関数のroundを使って必要がない桁は切り捨ててください。

## これからやりたいこと
- 罫線を自動で表現。
  - 今のところ、手動で`\hline`で設定が必要。
- 表のタイトルとlabelを設定。
  - 同じく、今のところは手動で追加が必要。
