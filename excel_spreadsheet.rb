require 'win32ole'

class ExcelSpreadsheet
  def initialize
    @excel = WIN32OLE::new('excel.Application')
    @workbook = @excel.Workbooks.Add

  end

  def add_worksheet(name)
    new_worksheet = ExcelWorksheet.new(@workbook.Worksheets.Add, name)

    # Remove the default Sheet1, Sheet2, and Sheet3 worksheets that Excel automatically creates
    @workbook.worksheets.each { |ws| ws.delete if ws.name =~ /^Sheet\d$/ }

    new_worksheet
  end

  def save_and_close(directory, filename_prefix)
    filename = File.join(directory, filename_prefix + Time.now.strftime('%Y-%m-%d_%H-%M') + '.xlsx')
    @workbook.saveAs(filename)

    @workbook.close
    @excel.DisplayAlerts = false
    @excel.workbooks.each do |wb|
      wb.close
    end
    @excel.quit
  end
end

class ExcelWorksheet
  def initialize(worksheet, name)
    @worksheet = worksheet
    @worksheet.name = name
    @worksheet.activate

    @current_row_number = 1
  end

  def add_row(data)
    current_column_letter = 'A'
    data.each do |item|
      @worksheet.range(current_column_letter + @current_row_number.to_s).Value = item
      current_column_letter = (current_column_letter.ord + 1).chr
    end
    @current_row_number = @current_row_number + 1
  end
end