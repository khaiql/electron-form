require 'axlsx'
require 'byebug'

result_path = File.expand_path('../results.xlsx', __FILE__)
FileUtils.rm(result_path) if File.exists?(result_path)

@customer_name = "Khai Le"
@customer_business = ""
@customer_street = "Bogenbay Batyra 134"
@customer_city = "Almaty"
@customer_postcode = "050000"
@customer_country_code = "KZ"

@external_item_id_barcode = "IC280000016KZ"
@external_item_id = "IC280000016KZ"

@package_items = [
  {description: "T-Shirt of Cotton", quantity: 2, value: 20.00},
  {description: "T-Shirt of Other Textile Materials", quantity: 1, value: 10.00},
]
@total_value = 30.00


def add_logo(sheet)
  logo = File.expand_path('images/kaz_logo.png')
  sheet.add_image(image_src: logo, noSelect: false, noMove: false) do |image|
    image.height = 22
    image.width = 130
    image.start_at 0, 0
  end
end

def add_signature(sheet, start_at_row, start_at_col)
  signature = File.expand_path('images/signature.png')
  sheet.add_image(image_src: signature) do |image|
    image.height = 62
    image.width = 82
    image.start_at start_at_row, start_at_col
  end
end

def add_cell(sheet, range, value, formatting = {}, style = nil)
  cell = sheet[range]
  raise "Must add row first" if cell.nil?
  cell.value = value
  formatting ||= {}
  cell.b = !!formatting[:bold]
  cell.sz = formatting[:size] if formatting[:size] && formatting[:size].to_i > 0
  cell.style = style if style
  if value.is_a?(Axlsx::RichText)
    cell.type = :richtext
    value.cell = cell
  end
  cell
end

def add_packages_items(sheet, start_row, package_items, style = 0)
  package_items.each_with_index do |package_item, index|
    row = start_row + index
    sheet.merge_cells("B#{row}:F#{row}")
    add_cell(sheet, "B#{row}", package_item[:description])
    add_cell(sheet, "G#{row}", package_item[:quantity], {}, style)
    sheet.merge_cells("J#{row}:K#{row}")
    add_cell(sheet, "J#{row}", package_item[:value], {}, style)
  end
end

def add_left_border(cell)

end

def get_cell_style(sheet, font_styles = nil, alignment = nil, borders = [])
  font_styles ||= {}

  if borders.length > 0
    border_style = { :style => :thin, :color => "00", edges: borders }
  end
  sheet.styles.add_style b: !!font_styles[:bold],
                        sz: font_styles[:size],
                        border: border_style,
                        alignment: alignment
end

def apply_style(sheet, cell_identifer, style)
  cells = sheet[cell_identifer]
  if cells.instance_of?(Array)
    cells.each { |cell| cell.style = style}
  else
    cells.style = style
  end
end

def add_total_value_row(sheet, row, total_value, style = 0)
  sheet.merge_cells("J#{row}:J#{row+1}")
  add_cell(sheet, "J#{row}", "Total value (6)\nОбщая стоимость", {size: 8})

  sheet.merge_cells("K#{row}:K#{row+1}")
  add_cell(sheet, "K#{row}", total_value, {size: 9}, style)
end

p = Axlsx::Package.new
wb = p.workbook

wb.styles do |s|
  # centered_cell = s.add_style({:alignment => {:horizontal => :center, :vertical => :center, :wrap_text => true}})
  top_wrapped_text_cell = s.add_style({:alignment => {:vertical => :top, :wrap_text => true}})
  center_left_wrapped_text = s.add_style({:alignment => {:vertical => :center, :wrap_text => true}})
  # borders
  surrounding_border = s.add_style(:border => Axlsx::STYLE_THIN_BORDER)
  left_top_border = s.add_style(border: {style: :thin, :color => '000000', edges: [:left, :top]})
  top_right_border = s.add_style(border: {style: :thin, :color => '000000', edges: [:top, :right]})
  right_bottom_border = s.add_style(border: {style: :thin, :color => '000000', edges: [:right, :bottom]})
  left_bottom_border = s.add_style(border: {style: :thin, :color => '000000', edges: [:left, :bottom]})
  left_border = s.add_style(border: {style: :thin, :color => '000000', edges: [:left]})
  top_border = s.add_style(border: {style: :thin, :color => '000000', edges: [:top]})
  right_border = s.add_style(border: {style: :thin, :color => '000000', edges: [:right]})
  bottom_border = s.add_style(border: {style: :thin, :color => '000000', edges: [:bottom]})

  street_cell_style = s.add_style(:alignment => {:vertical => :center, :wrap_text => true})
  # left_border_and_alignment_style = s.add_style(border: {style: :thin, :color => '000000', edges: [:left]}, :alignment => {:vertical => :center, :wrap_text => true})

  wb.add_worksheet(name: 'Таможенная декларация') do |sheet|
    center_alignment = {:horizontal => :center, :vertical => :center, :wrap_text => true}
    centered_cell = get_cell_style(sheet, {}, center_alignment)
    left_border_and_alignment_style = get_cell_style(sheet, {}, {:vertical => :center, :wrap_text => true}, [:left] )
    (0..79).each do |row|
      sheet.add_row [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil]
    end



    # sheet.sheet_view.show_grid_lines = false

    sheet.merge_cells("O1:O2")
    add_cell(sheet, "O1", "CN23", {bold: true, size: 16}, centered_cell)

    sheet.merge_cells("G1:I2")
    add_cell(sheet, "G1", " =CONSIGNMENT= ", {bold: true, size: 14}, centered_cell)

    sheet.merge_cells("A1:B2")
    add_logo(sheet)

    add_cell(sheet, "K1", "CUSTOMS DECLARATION")
    add_cell(sheet, "K2", "ТАМОЖЕННАЯ ДЕКЛАРАЦИЯ", {bold: true})

    sheet.merge_cells("A3:A5")
    add_cell(sheet, "A3", "From:", {bold: true, size: 7}, surrounding_border)

    add_cell(sheet, "B3", "Name", {size: 7}, top_border)
    add_cell(sheet, "B4", "Фамилия", {}, left_bottom_border)
    add_cell(sheet, "C3", "", {}, top_right_border)
    add_cell(sheet, "C4", "", {}, right_bottom_border)
    add_cell(sheet, "B5", "Business", {size: 7}, left_border)
    add_cell(sheet, "B6", "Компания/Фирма", {}, left_bottom_border)
    add_cell(sheet, "C5", "", {}, right_border)
    add_cell(sheet, "C6", "", {}, right_bottom_border)

    sheet.merge_cells("D3:G4")
    add_cell(sheet, "D3", "Kazpost GmbH", {bold: true}, get_cell_style(sheet, nil, center_alignment, [:top, :right]))
    sheet.merge_cells("D5:G6")
    add_cell(sheet, "D5", "Acc. Contact No.____________", {bold: true}, centered_cell)
    apply_style(sheet, "D6:F6", get_cell_style(sheet, nil, center_alignment, [:bottom]))
    apply_style(sheet, "G6", get_cell_style(sheet, nil, center_alignment, [:right, :bottom]))
    apply_style(sheet, "G5", get_cell_style(sheet, nil, center_alignment, [:right]))
    apply_style(sheet, "D3:F3", get_cell_style(sheet, nil, center_alignment, [:top]))
    apply_style(sheet, "G3", get_cell_style(sheet, nil, center_alignment, [:top, :right]))
    apply_style(sheet, "G4", get_cell_style(sheet, nil, center_alignment, [:right]))

    sheet.merge_cells("H3:I3")
    add_cell(sheet, "H3", "Importer's reference (tax code/VAT No./importer code) (optional)", {size: 6})
    sheet.merge_cells("H4:I4")
    add_cell(sheet, "H4", "Таможенная ссылка отправителя (если имеется)", {size: 6}, top_wrapped_text_cell)
    apply_style(sheet, "H3:I3", get_cell_style(sheet, nil, {vertical: :center, horizontal: :left, wrap_text: true}, [:top]))
    apply_style(sheet, "I3", get_cell_style(sheet, nil, {}, [:top, :right]))
    apply_style(sheet, "I4:I6", get_cell_style(sheet, nil, {}, [:right]))
    apply_style(sheet, "H6:I6", get_cell_style(sheet, nil, {}, [:bottom]))
    apply_style(sheet, "I6", get_cell_style(sheet, nil, {}, [:bottom, :right]))

    add_cell(sheet, "A7", "Из", {bold: true, size: 7}, top_border)
    add_cell(sheet, "B7", "Street", {size: 7}, left_border)
    add_cell(sheet, "C7", "")
    add_cell(sheet, "B8", "Улица", {}, left_bottom_border)
    add_cell(sheet, "C8", "", nil, bottom_border)

    sheet.merge_cells("D7:I8")
    add_cell(sheet, "D7", "Carl-Zeiss-Str., 25, (+49 5131 502 9504)", {bold: true}, street_cell_style)
    apply_style(sheet, "D8:H8", get_cell_style(sheet, nil, nil, [:bottom]))
    apply_style(sheet, "I8", get_cell_style(sheet, nil, nil, [:bottom, :right]))
    apply_style(sheet, "I7", get_cell_style(sheet, nil, nil, [:right]))

    add_cell(sheet, "B9", "Postcode", {size: 7}, left_border)
    add_cell(sheet, "B10", "Почтовый индекс", {}, left_border)
    apply_style(sheet, "B10", get_cell_style(sheet, nil, nil, [:left]))
    add_cell(sheet, "C9", "", {}, top_border)
    add_cell(sheet, "C10", "", {}, bottom_border)
    sheet.merge_cells("D9:E10")
    add_cell(sheet, "D9", "30827", {bold: true}, center_left_wrapped_text)
    apply_style(sheet, "B10:E10", get_cell_style(sheet, nil, nil, [:bottom]))

    add_cell(sheet, "F9", "City", {size: 7}, left_border)
    add_cell(sheet, "F10", "Город", {}, bottom_border)
    sheet.merge_cells("G9:I10")
    add_cell(sheet, "G9", "Garbsen, Niedersachsen", {bold: true}, center_left_wrapped_text)
    apply_style(sheet, "G10:I10", get_cell_style(sheet, nil, nil, [:bottom]))
    apply_style(sheet, "I10", get_cell_style(sheet, nil, nil, [:bottom, :right]))
    apply_style(sheet, "I9", get_cell_style(sheet, nil, nil, [:right]))

    add_cell(sheet, "B11", "Country", {size: 7}, left_top_border)
    add_cell(sheet, "B12", "Страна", {}, left_bottom_border)
    add_cell(sheet, "C11", "")
    add_cell(sheet, "C12", "", nil, bottom_border)
    sheet.merge_cells("D11:I12")
    add_cell(sheet, "D11", "DE", {bold: true}, center_left_wrapped_text)
    apply_style(sheet, "D12:H12", get_cell_style(sheet, nil, nil, [:bottom]))
    apply_style(sheet, "I12", get_cell_style(sheet, nil, nil, [:bottom, :right]))
    apply_style(sheet, "I11", get_cell_style(sheet, nil, nil, [:right]))

    add_cell(sheet, "A13", "To:", {bold: true, size: 8}, surrounding_border)
    add_cell(sheet, "A14", "B:", {bold: true, size: 8}, surrounding_border)

    add_cell(sheet, "B13", "Name", {size: 7})
    add_cell(sheet, "B14", "Фамилия")
    add_cell(sheet, "C13", "")
    add_cell(sheet, "C14", "", nil, bottom_border)
    sheet.merge_cells("D13:I14")
    add_cell(sheet, "D13", @customer_name, {bold: true}, center_left_wrapped_text)
    apply_style(sheet, "D14:H14", get_cell_style(sheet, nil, nil, [:bottom]))
    apply_style(sheet, "I14", get_cell_style(sheet, nil, nil, [:bottom, :right]))
    apply_style(sheet, "I13", get_cell_style(sheet, nil, nil, [:right]))

    add_cell(sheet, "B15", "Business", {size: 7}, left_top_border)
    add_cell(sheet, "B16", "Компания/Фирма", nil, left_bottom_border)
    add_cell(sheet, "C15", "", nil, top_border)
    add_cell(sheet, "C16", "", nil, bottom_border)
    sheet.merge_cells("D15:I16")
    add_cell(sheet, "D15", @customer_business, {bold: true}, center_left_wrapped_text)
    apply_style(sheet, "D16:H16", get_cell_style(sheet, nil, nil, [:bottom]))
    apply_style(sheet, "I16", get_cell_style(sheet, nil, nil, [:bottom, :right]))
    apply_style(sheet, "I15", get_cell_style(sheet, nil, nil, [:right]))

    add_cell(sheet, "B17", "Street", {size: 7}, left_border)
    add_cell(sheet, "B18", "Улица", nil, left_bottom_border)
    add_cell(sheet, "C17", "")
    add_cell(sheet, "C18", "", nil, bottom_border)
    sheet.merge_cells("D17:I18")
    add_cell(sheet, "D17", @customer_street, {bold: true}, center_left_wrapped_text)
    apply_style(sheet, "D18:H18", get_cell_style(sheet, nil, nil, [:bottom]))
    apply_style(sheet, "I18", get_cell_style(sheet, nil, nil, [:bottom, :right]))
    apply_style(sheet, "I16", get_cell_style(sheet, nil, nil, [:right]))

    add_cell(sheet, "B19", "Postcode", {size: 7}, left_border)
    add_cell(sheet, "B20", "Почтовый индекс", nil, left_bottom_border)
    add_cell(sheet, "C19", "")
    add_cell(sheet, "C20", "", nil, bottom_border)
    sheet.merge_cells("D19:E20")
    add_cell(sheet, "D19", @customer_postcode, {bold: true}, center_left_wrapped_text)

    add_cell(sheet, "F19", "City", {size: 7}, left_border)
    add_cell(sheet, "F20", "Город", nil, left_bottom_border)
    apply_style(sheet, "D20:E20", get_cell_style(sheet, nil, nil, [:bottom]))
    sheet.merge_cells("G19:I20")
    add_cell(sheet, "G19", @customer_city, {bold: true}, center_left_wrapped_text)
    apply_style(sheet, "G20:H20", get_cell_style(sheet, nil, nil, [:bottom]))
    apply_style(sheet, "I20", get_cell_style(sheet, nil, nil, [:bottom, :right]))
    apply_style(sheet, "I19", get_cell_style(sheet, nil, nil, [:right]))
    apply_style(sheet, "I17", get_cell_style(sheet, nil, nil, [:top, :right]))
    apply_style(sheet, "B10", get_cell_style(sheet, nil, nil, [:left]))

    add_cell(sheet, "B21", "Country", {size: 7}, left_border)
    add_cell(sheet, "B22", "Страна", nil, left_bottom_border)
    add_cell(sheet, "C21", "")
    add_cell(sheet, "C22", "", nil, bottom_border)
    sheet.merge_cells("D21:I22")
    add_cell(sheet, "D21", @customer_country_code, {bold: true}, center_left_wrapped_text)
    apply_style(sheet, "D22:H22", get_cell_style(sheet, nil, nil, [:bottom, :top]))
    apply_style(sheet, "I22", get_cell_style(sheet, nil, nil, [:bottom, :right]))
    apply_style(sheet, "I21", get_cell_style(sheet, nil, nil, [:top, :right]))

    sheet.merge_cells("J4:L4")
    add_cell(sheet, "J4", "No. of item (barcode, if any)", {bold: true}, center_left_wrapped_text)

    sheet.merge_cells("J5:L5")
    add_cell(sheet, "J5", "№ отправления (штриховой код,", {bold: true}, center_left_wrapped_text)

    sheet.merge_cells("J6:L6")
    add_cell(sheet, "J6", "если имеется)", {bold: true}, center_left_wrapped_text)

    sheet.merge_cells("M4:N4")
    add_cell(sheet, "M4", "May be opened officially", {bold: true}, left_border_and_alignment_style)

    sheet.merge_cells("M5:N5")
    add_cell(sheet, "M5", "Может быть вскрыто в служебном", {bold: true}, left_border_and_alignment_style)

    sheet.merge_cells("M6:N6")
    add_cell(sheet, "M6", "порядке", {bold: true}, left_border_and_alignment_style)


    sheet.merge_cells("J7:O14")
    add_cell(sheet, "J7", @external_item_id_barcode, {bold: true, size: 30}, centered_cell)

    sheet.merge_cells("J15:O16")
    add_cell(sheet, "J15", @external_item_id, {bold: true}, centered_cell)

    sheet.merge_cells("J17:O19")
    add_cell(sheet, "J17", "Importer's reference (if any) (taxcode/VAT No./importer code) (optional)\x0D\x0AРеквизиты импортера (если имеются) (ИНН/№НДС/индекс импортера) (факультативно)", {size: 7}, top_wrapped_text_cell)

    sheet.merge_cells("J20:O22")
    add_cell(sheet, "J20", "Importer's telephone/fax/e-mail (if known)\x0D\x0A№ телефона/факса/e-mail импортера (если известен)", {size: 7}, top_wrapped_text_cell)

    # details table header
    sheet.merge_cells("B23:F23")
    add_cell(sheet, "B23", "Detailed description  of contents (1)", {size: 7}, left_top_border)
    sheet.merge_cells("B24:F24")
    add_cell(sheet, "B24", "Подробное описание вложения (1)", {size: 7}, left_bottom_border)

    add_cell(sheet, "G23", "Quantity (2) ", {size: 7}, left_top_border)
    add_cell(sheet, "G24", "Количество (2) ", {size: 7}, left_bottom_border)

    sheet.merge_cells("H23:I23")
    add_cell(sheet, "H23", "Net weight (in kg) (3) ", {size: 7}, left_top_border)
    sheet.merge_cells("H24:I24")
    add_cell(sheet, "H24", "Вес нетто (в кг) (3) ", {size: 7}, left_bottom_border)

    sheet.merge_cells("J23:K23")
    add_cell(sheet, "J23", "Value (5) ", {size: 7}, left_top_border)
    sheet.merge_cells("J24:K24")
    add_cell(sheet, "J24", "Стоимость (5) ", {size: 7}, left_bottom_border)

    sheet.merge_cells("L23:O23")
    add_cell(sheet, "L23", "For commercial items only\nТолько для коммерческих отправлений", {size: 7}, top_wrapped_text_cell)

    sheet.merge_cells("L24:M24")
    add_cell(sheet, "L24", "HS tariff number (7)\nКод ТНВЭД (7)", {size: 7}, top_wrapped_text_cell)

    sheet.merge_cells("N24:O24")
    add_cell(sheet, "N24", "Country of origin of goods (8)\nСтрана происхождения товаров(8)", {size: 7}, top_wrapped_text_cell)

    @package_items.concat([{}, {}, {}])
    add_packages_items(sheet, 25, @package_items, centered_cell)
    add_total_value_row(sheet, 25 + @package_items.length, @total_value, centered_cell)

    total_row_index = 25 + @package_items.length

    sheet.merge_cells("H#{total_row_index}:H#{total_row_index + 1}")
    add_cell(sheet, "H#{total_row_index}", "Total gross weight (4)", {size: 7}, center_left_wrapped_text)
    sheet.merge_cells("I#{total_row_index}:I#{total_row_index+1}")
    add_cell(sheet, "I#{total_row_index}", '')

    sheet.merge_cells("J#{total_row_index}:J#{total_row_index+1}")
    add_cell(sheet, "J#{total_row_index}", "Total value (6)\nОбщая стоимость", {size: 8}, center_left_wrapped_text)

    sheet.merge_cells("K#{total_row_index}:K#{total_row_index+1}")
    add_cell(sheet, "K#{total_row_index}", @total_value, {size: 9}, centered_cell)

    sheet.merge_cells("L#{total_row_index}:M#{total_row_index}")
    add_cell(sheet, "L#{total_row_index}", "Postal charges/fees (9)", {size: 7}, center_left_wrapped_text)
    sheet.merge_cells("L#{total_row_index+1}:M#{total_row_index+1}")
    add_cell(sheet, "L#{total_row_index+1}", "Почтовые сборы/Расходы (9)", {size: 7}, center_left_wrapped_text)

    footer_row_index = total_row_index + 2
    add_cell(sheet, "B#{footer_row_index}", "Category of item (10)", {size: 7})
    add_cell(sheet, "B#{footer_row_index + 1}", "Категория отправления (10)")
    sheet.merge_cells("E#{footer_row_index}:E#{footer_row_index + 1}")
    add_cell(sheet, "F#{footer_row_index}", "Commercial sample", {size: 7})
    add_cell(sheet, "F#{footer_row_index + 1}", "Коммерческий образец")
    add_cell(sheet, "H#{footer_row_index}", "Explanation:", {size: 7})
    add_cell(sheet, "H#{footer_row_index + 1}", "Пояснение :")
    sheet.merge_cells("B#{footer_row_index + 2}:B#{footer_row_index + 3}")
    add_cell(sheet, "C#{footer_row_index + 2}", "Gift", {size: 7})
    add_cell(sheet, "C#{footer_row_index + 3}", "Подарок")
    sheet.merge_cells("E#{footer_row_index + 2}:E#{footer_row_index+3}")
    add_cell(sheet, "F#{footer_row_index + 2}", "Returned goods", {size: 7})
    add_cell(sheet, "F#{footer_row_index + 3}", "Возврат товара")
    sheet.merge_cells("B#{footer_row_index + 4}:B#{footer_row_index + 5}")
    add_cell(sheet, "C#{footer_row_index + 4}", " Documents", {size: 7})
    add_cell(sheet, "C#{footer_row_index + 5}", " Документ")
    sheet.merge_cells("E#{footer_row_index + 4}:E#{footer_row_index+5}")
    add_cell(sheet, "E#{footer_row_index + 4}", "X", {bold: true, size: 16}, centered_cell)
    add_cell(sheet, "F#{footer_row_index + 4}", "Other", {size: 7})
    add_cell(sheet, "F#{footer_row_index + 5}", "Прочее")

    sheet.merge_cells("B#{footer_row_index + 6}:K#{footer_row_index + 6}")
    add_cell(sheet, "B#{footer_row_index + 6}", "Comments (11): (e.g.: goodssubject to quarantine, sanitary/phytosanitary inspection or other instructions)", {size: 9})
    sheet.merge_cells("B#{footer_row_index + 7}:K#{footer_row_index + 7}")
    add_cell(sheet, "B#{footer_row_index + 7}", "Примечания (11): (напр., товар, подлежащий карантину/санитарному, фитосанитарному контролю или попадающий под другие ограничения)", {size: 9})
    sheet.merge_cells("B#{footer_row_index + 8}:K#{footer_row_index + 8}")

    sheet.merge_cells("L#{footer_row_index}:O#{footer_row_index}")
    add_cell(sheet, "L#{footer_row_index}", "Office of origin/Date of posting", {size: 8})
    sheet.merge_cells("L#{footer_row_index + 1}:O#{footer_row_index + 1}")
    add_cell(sheet, "L#{footer_row_index + 1}", "Учреждение подачи/Дата подачи", {size: 8})
    sheet.merge_cells("L#{footer_row_index + 2}:O#{footer_row_index + 2}")
    sheet.merge_cells("L#{footer_row_index + 3}:O#{footer_row_index + 3}")
    add_cell(sheet, "L#{footer_row_index + 3}", "HAMBURG E KAZPOST", {bold: true, size: 8}, centered_cell)
    sheet.merge_cells("L#{footer_row_index + 4}:O#{footer_row_index + 4}")
    add_cell(sheet, "L#{footer_row_index + 4}", Date.today.strftime("%d.%m.%Y"), {bold: true, size: 8}, centered_cell)
    sheet.merge_cells("L#{footer_row_index + 5}:O#{footer_row_index + 8}")

    sheet.merge_cells("B#{footer_row_index + 9}:K#{footer_row_index + 9}")
    sheet.merge_cells("B#{footer_row_index + 10}:B#{footer_row_index + 11}")
    add_cell(sheet, "C#{footer_row_index + 10}", "Licence (12)", {size: 7})
    add_cell(sheet, "C#{footer_row_index + 11}", "Лицензия (12)")
    add_cell(sheet, "B#{footer_row_index + 12}", "No(s). Of licence(s)/№ лицензии (-ий)", {size: 7})

    sheet.merge_cells("E#{footer_row_index + 10}:E#{footer_row_index+11}")
    add_cell(sheet, "F#{footer_row_index + 10}", "Certificate (13)", {size: 7})
    add_cell(sheet, "F#{footer_row_index + 11}", "Сертификат (13)")
    add_cell(sheet, "E#{footer_row_index + 12}", "No(s). Of certificate(s)/№ сертификата", {size: 7})

    sheet.merge_cells("H#{footer_row_index + 10}:H#{footer_row_index+11}")
    add_cell(sheet, "I#{footer_row_index + 10}", "Invoice  (14)", {size: 7})
    add_cell(sheet, "I#{footer_row_index + 11}", "Счет (14)")
    add_cell(sheet, "H#{footer_row_index + 12}", "No. of invoice/№ счета", {size: 7})

    sheet.merge_cells("L#{footer_row_index + 9}:O#{footer_row_index+9}")
    add_cell(sheet, "L#{footer_row_index+9}", "I certify that the particulars given in this customs declaration are correct and that this item does not contain any dangerous articles prohibited by legislation or by postal or customs regulations", {size: 7}, top_wrapped_text_cell)
    sheet.merge_cells("L#{footer_row_index + 10}:O#{footer_row_index+10}")
    add_cell(sheet, "L#{footer_row_index+10}", "Я подтверждаю, что указанные в настоящей таможенной декларации сведения являются достоверными, и что в этом отправлении не содержится никаких опасных или запрещенных законодательством или почтовой или таможенной регламентацией предметов ", {size: 7}, top_wrapped_text_cell)
    sheet.merge_cells("L#{footer_row_index + 11}:M#{footer_row_index+11}")
    add_cell(sheet, "L#{footer_row_index+11}", "Дата и подпись отправителя (15)", {size: 7}, center_left_wrapped_text)
    sheet.merge_cells("L#{footer_row_index + 12}:M#{footer_row_index+12}")
    add_cell(sheet, "L#{footer_row_index+12}", Date.today.strftime("%d.%m.%Y"), {size: 10}, center_left_wrapped_text)

    sheet.merge_cells("O#{footer_row_index + 11}:P#{footer_row_index+11}")
    add_cell(sheet, "O#{footer_row_index+11}", "Date and sender’s signature (15)", {size: 7}, center_left_wrapped_text)

    sheet.merge_cells("O#{footer_row_index + 12}:P#{footer_row_index+12}")
    cell = sheet["O#{footer_row_index + 12}"]
    add_signature(sheet, *cell.pos)

    sheet.sheet_view.show_grid_lines = false
    sheet.page_setup.fit_to :width => 1, :height => 1
    sheet.column_widths 4, 4, 11, 13, 5, 7, 13, 17, 7, 13, 8.5, 19, 10, 10, 14.5, 4.5, nil
  end
end



p.serialize result_path