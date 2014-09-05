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

end
