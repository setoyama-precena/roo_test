require 'roo'

class RooController < ApplicationController

  def post_xlsx
    file = params[:xlsx_file]
    xlsx = Roo::Excelx.new(file.path, nil, :ignore)
    @list_title = xlsx.cell('B', 3)
    @date_sample = xlsx.cell('C', 4)
    @windows_char_sample = xlsx.cell('C', 6)
    @special_char_sample1 = xlsx.cell('C', 8)
    @special_char_sample2 = xlsx.cell('C', 10)

    @sheet2_list_tilte = xlsx.sheet('シート2').cell('B', 2)

  end

  def show
    html = render_to_string template: "roo/show"

    # PDFKitを作成
    pdf = PDFKit.new(html, encoding: "UTF-8")

    # スタイルシートの設定
    pdf.stylesheets << "#{Rails.root}/app/assets/stylesheets/roo.css.scss"

    # 画面にPDFを表示する
    # to_pdfメソッドでPDFファイルに変換する
    # 他には、to_fileメソッドでPDFファイルを作成できる
    # disposition: "inline" によりPDFはダウンロードではなく画面に表示される
    send_data pdf.to_pdf,
              filename:    "test.pdf",
              type:        "application/pdf"
  end

end
