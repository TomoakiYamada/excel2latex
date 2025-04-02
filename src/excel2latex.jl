# set your path
cd("/Users/tomoakiyamada/Documents/Working/excel2latex/src/")

using XLSX

"""
罫線スタイルを表現する構造体
"""
struct BorderStyle
    top::Symbol
    bottom::Symbol
    left::Symbol
    right::Symbol
end

"""
Excel表をLaTeX形式に変換する構造体
"""
struct TableConverter
    filename::String
    sheet::String
    xf::XLSX.XLSXFile
    sheet_data::XLSX.Worksheet
    first_cell::String
    last_cell::String
end

"""
コンバーター初期化
"""
function TableConverter(filename::String, sheet::String="Sheet1")
    xf = XLSX.readxlsx(filename)
    sheet_data = xf[sheet]

    rng = XLSX.get_dimension(sheet_data)
    first_cell = XLSX.encode_column_number(rng.start.column_number) * string(rng.start.row_number)
    last_cell = XLSX.encode_column_number(rng.stop.column_number) * string(rng.stop.row_number)

    return TableConverter(filename, sheet, xf, sheet_data, first_cell, last_cell)
end

"""
罫線情報を取得 (新しいAPIに対応)
"""
function get_cell_borders(tc::TableConverter, row::Int, col::Int)
    # 現在のXLSX.jlでは直接的な罫線情報の取得方法が変更されている
    # とりあえずデフォルト値を返す
    return BorderStyle(:none, :none, :none, :none)

    # TODO: 実際の罫線情報の取得方法を実装
    # XLSX.jlのドキュメントやソースコードを確認して
    # 適切なAPIを見つける必要がある
end

"""
LaTeX表の生成
"""
function generate_latex_table(tc::TableConverter)
    latex = "\\begin{table}[htbp]\n\\centering\n\\begin{tabular}"

    # データの取得
    data = tc.sheet_data[tc.first_cell * ":" * tc.last_cell]

    # 列数を取得して列指定を生成
    num_cols = size(data, 2)
    latex *= "{" * repeat("c", num_cols) * "}\n"

    # データの処理
    for row in 1:size(data, 1)
        row_data = String[]
        for col in 1:size(data, 2)
            cell_value = data[row, col]
            cell_value = isnothing(cell_value) ? "" : string(cell_value)
            push!(row_data, cell_value)

            # 罫線情報の取得（現在は実装待ち）
            borders = get_cell_borders(tc, row, col)
        end

        latex *= join(row_data, " & ") * " \\\\\n"
    end

    latex *= "\\end{tabular}\n\\end{table}"
    return latex
end

"""
メイン処理
"""
function convert_excel_to_latex(filename::String, sheet::String="Sheet1")
    tc = TableConverter(filename, sheet)
    latex_table = generate_latex_table(tc)
    return latex_table
end

function main()
    latex_table = convert_excel_to_latex("../excel/sample.xlsx")
    println(latex_table)

    write("../output/table.tex", latex_table)
end

main()
