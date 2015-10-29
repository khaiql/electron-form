require './form'

class FormCP71 < Form

  def add_total_value_row(sheet, row, total_value, style = 0)
    sheet.merge_cells("J#{row}:J#{row+1}")
    add_cell(sheet, "J#{row}", "Total value (6)\nОбщая стоимость", {size: 8})

    sheet.merge_cells("K#{row}:K#{row+1}")
    add_cell(sheet, "K#{row}", total_value, {size: 9}, style)
  end

  def add_packages_items(sheet, start_row, package_items, style = 0)
    package_items.each_with_index do |package_item, index|
      row = start_row + index
      sheet.merge_cells("B#{row}:F#{row}")
      add_cell(sheet, "B#{row}", package_item[:description])
      add_cell(sheet, "G#{row}", package_item[:quantity])
      sheet.merge_cells("J#{row}:K#{row}")
      add_cell(sheet, "J#{row}", package_item[:value])
      apply_style(sheet, "B#{row}", get_cell_style(sheet, nil, nil, [:left, :bottom]))
      apply_style(sheet, "C#{row}:F#{row}", get_cell_style(sheet, nil, nil, [:bottom]))
      apply_style(sheet, "G#{row}", get_cell_style(sheet, nil, {vertical: :center, horizontal: :center}, [:left, :bottom]))
      apply_style(sheet, "H#{row}", get_cell_style(sheet, nil, nil, [:left, :bottom]))
      apply_style(sheet, "I#{row}", get_cell_style(sheet, nil, nil, [:bottom]))
      apply_style(sheet, "J#{row}", get_cell_style(sheet, nil, {vertical: :center, horizontal: :center}, [:left, :bottom]))
      apply_style(sheet, "K#{row}", get_cell_style(sheet, nil, nil, [:bottom]))
      apply_style(sheet, "L#{row}", get_cell_style(sheet, nil, nil, [:left, :bottom]))
      apply_style(sheet, "M#{row}", get_cell_style(sheet, nil, nil, [:bottom]))
      apply_style(sheet, "N#{row}", get_cell_style(sheet, nil, nil, [:left, :bottom]))
      apply_style(sheet, "O#{row}", get_cell_style(sheet, nil, nil, [:bottom, :right]))
    end
  end

  def open_file
    cmd = "open #{self.form_name}"
    value = %x[ #{cmd} ]
  end

  def process
    p = Axlsx::Package.new
    wb = p.workbook

    wb.styles do |s|
      # centered_cell = s.add_style({:alignment => {:horizontal => :center, :vertical => :center, :wrap_text => true}})
      top_wrapped_text_cell = s.add_style({:alignment => {:vertical => :top, :wrap_text => true}})
      center_left_wrapped_text = s.add_style({:alignment => {:vertical => :center, :wrap_text => true}})
      # borders
      # surrounding_border = s.add_style(:border => Axlsx::STYLE_THIN_BORDER)
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
        (0..73).each do |row|
          sheet.add_row [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil]
        end



        # sheet.sheet_view.show_grid_lines = false

        # AAAAAAAAAA
        sheet.merge_cells("A1:C2")
        add_logo(sheet, 0, 0)
        add_cell(sheet, "A3", "Из", {}, top_right_border)

        sheet.merge_cells("A4:A6")
        add_cell(sheet, "A4", "From:", {bold: true, size: 7})

        sheet.merge_cells("A14:A15")
        add_cell(sheet, "A14", "To", {bold: true, size: 8}, top_right_border)

        # BBBBBBBBBB

        sheet.merge_cells("B3:C3")
        add_cell(sheet, "B3", "Фамилия", {size: 7}, top_border)
        add_cell(sheet, "B4", "Name", {}, left_bottom_border)
        add_cell(sheet, "B5", "Business", {}, left_border)
        add_cell(sheet, "B6", "Компания/Фирма", {}, left_bottom_border)


        sheet.merge_cells("B5:C5")
        add_cell(sheet, "B5", "Компания/Фирма", {}, left_bottom_border)

        sheet.merge_cells("B6:C6")
        add_cell(sheet, "B6", "Business", {}, left_bottom_border)

        sheet.merge_cells("B7:C7")
        add_cell(sheet, "B7", "Улица", {}, left_bottom_border)

        add_cell(sheet, "B8", "Street", {}, left_bottom_border)

        sheet.merge_cells("B9:C9")
        add_cell(sheet, "B9", "Почтовый индекс", {}, left_bottom_border)

        sheet.merge_cells("B10:C10")
        add_cell(sheet, "B10", "Postcode", {}, left_bottom_border)

        sheet.merge_cells("B11:C11")
        add_cell(sheet, "B11", "Страна", {}, left_bottom_border)

        add_cell(sheet, "B12", "Country", {}, left_bottom_border)

        sheet.merge_cells("B13:C13")
        add_cell(sheet, "B13", "Фамилия", {}, left_bottom_border)

        add_cell(sheet, "B14", "Name", {}, left_bottom_border)

        sheet.merge_cells("B15:C15")
        add_cell(sheet, "B15", "Компания/Фирма", {}, left_bottom_border)

        sheet.merge_cells("B16:C16")
        add_cell(sheet, "B16", "Business", {}, left_bottom_border)

        sheet.merge_cells("B17:C17")
        add_cell(sheet, "B17", "Улица", {}, left_bottom_border)

        add_cell(sheet, "B18", "Street", {}, left_bottom_border)

        sheet.merge_cells("B19:C19")
        add_cell(sheet, "B19", "Почтовый индекс", {}, left_bottom_border)

        sheet.merge_cells("B20:C20")
        add_cell(sheet, "B20", "Postcode", {}, left_bottom_border)

        sheet.merge_cells("B21:C21")
        add_cell(sheet, "B21", "Страна", {}, left_bottom_border)

        add_cell(sheet, "B22", "Country", {}, left_bottom_border)

        sheet.merge_cells("B23:E23")
        add_cell(sheet, "B23", "Учреждение обмена", {}, left_bottom_border)

        sheet.merge_cells("B24:C24")
        add_cell(sheet, "B24", "Office of exchange", {}, left_bottom_border)

        sheet.merge_cells("B27:E27")
        add_cell(sheet, "B27", "Указать требуемую услугу", {}, left_bottom_border)

        sheet.merge_cells("B28:E28")
        add_cell(sheet, "B28", "(зачеркнуть соответствующую клеточку)", {}, left_bottom_border)

        sheet.merge_cells("B29:E29")
        add_cell(sheet, "B29", "Please indicate service requiered", {}, left_bottom_border)

        sheet.merge_cells("B30:E30")
        add_cell(sheet, "B30", "(tick one box)", {}, left_bottom_border)

        sheet.merge_cells("B31:C31")
        add_cell(sheet, "B31", "Международное", {}, left_bottom_border)

        sheet.merge_cells("B32:C32")
        add_cell(sheet, "B32", "приоритетное", {}, left_bottom_border)

        sheet.merge_cells("B33:B34")
        add_cell(sheet, "B33", "X", {bold: true, size: 16}, centered_cell)

        sheet.merge_cells("B44:C44")
        add_cell(sheet, "B44", "Расписка", {}, left_bottom_border)

        sheet.merge_cells("B45:C45")
        add_cell(sheet, "B45", "получателя", {}, left_bottom_border)

        sheet.merge_cells("B46:C46")
        add_cell(sheet, "B46", "Declaration", {}, left_bottom_border)

        sheet.merge_cells("B47:C47")
        add_cell(sheet, "B47", "by addressee", {}, left_bottom_border)

        sheet.merge_cells("B48:G48")
        add_cell(sheet, "B48", "Подтверждаю, что сведения, указанные в настоящей таможенной", {}, left_bottom_border)

        sheet.merge_cells("B49:G49")
        add_cell(sheet, "B49", "декларации являются достоверными, и что в этом отправлении не", {}, left_bottom_border)

        sheet.merge_cells("B50:G50")
        add_cell(sheet, "B50", "содержится никаких опясных или запрещенных законодательством", {}, left_bottom_border)

        sheet.merge_cells("B51:G51")
        add_cell(sheet, "B51", "или почтовой или таможенной регламентацией предметов", {}, left_bottom_border)

        sheet.merge_cells("B53:E53")
        add_cell(sheet, "B53", "I certify that the particulars given in this customs", {}, left_bottom_border)

        sheet.merge_cells("B54:E54")
        add_cell(sheet, "B54", "declaration are correct and that this item does not", {}, left_bottom_border)

        sheet.merge_cells("B55:E55")
        add_cell(sheet, "B55", "contain any dangerous article prohibited by legislation", {}, left_bottom_border)

        sheet.merge_cells("B56:E56")
        add_cell(sheet, "B56", "or by postal or customs regulations", {}, left_bottom_border)


        # CCCCCCCCC
        add_cell(sheet, "C3", nil, {}, top_right_border)
        add_cell(sheet, "C4", nil, {}, right_bottom_border)
        add_cell(sheet, "C5", nil, {}, right_border)
        add_cell(sheet, "C6", nil, {}, right_bottom_border)

        add_cell(sheet, "C33", "Intemational", {}, left_bottom_border)
        add_cell(sheet, "C34", "Priority", {}, left_bottom_border)


        # DDDDDDDDD

        sheet.merge_cells("D3:G4")
        add_cell(sheet, "D3", "Kazpost GmbH", {bold: true}, get_cell_style(sheet, nil, center_alignment, [:top, :right]))

        sheet.merge_cells("D5:G6")
        add_cell(sheet, "D5", "Acc. Contact No.____________", {bold: true}, centered_cell)

        sheet.merge_cells("D7:I8")
        add_cell(sheet, "D7", "Carl-Zeiss-Str., 25, (+49 5131 502 9504)", {bold: true}, street_cell_style)

        sheet.merge_cells("D9:E10")
        add_cell(sheet, "D9", "30827", {bold: true}, center_left_wrapped_text)

        sheet.merge_cells("D11:I12")
        add_cell(sheet, "D11", "DE", {bold: true}, center_left_wrapped_text)

        sheet.merge_cells("D13:I14")
        add_cell(sheet, "D13", @customer_name, {bold: true}, center_left_wrapped_text)

        sheet.merge_cells("D17:I18")
        add_cell(sheet, "D17", @customer_street, {bold: true}, center_left_wrapped_text)

        sheet.merge_cells("D19:E20")
        add_cell(sheet, "D19", @customer_postcode, {bold: true}, center_left_wrapped_text)

        sheet.merge_cells("D21:I22")
        add_cell(sheet, "D21", @customer_country_code, {bold: true}, center_left_wrapped_text)

        sheet.merge_cells("D44:I44")
        add_cell(sheet, "D44", "Я получил посылку, описание которой в этом сопроводительном адресе", {}, left_bottom_border)

        sheet.merge_cells("D45:I45")
        add_cell(sheet, "D45", "Дата и подпись получателя", {}, left_bottom_border)

        sheet.merge_cells("D46:I46")
        add_cell(sheet, "D46", "I have received the parcel described on the notr", {}, left_bottom_border)

        sheet.merge_cells("D47:I47")
        add_cell(sheet, "D46", "Date and addressee's signature", {}, left_bottom_border)

        # EEEEEEEEE
        sheet.merge_cells("E1:I1")
        add_cell(sheet, "E1", "Отправление/посылка может быть вскрыто в служебном порядке", {bold: true, size: 14})

        sheet.merge_cells("E2:I2")
        add_cell(sheet, "E2", "May be opened officially", {bold: true, size: 14})

        add_cell(sheet, "E31", "Intemational", {}, left_bottom_border)
        add_cell(sheet, "E32", "эконом класса", {}, left_bottom_border)

        add_cell(sheet, "E33", "Intemational", {}, left_bottom_border)
        add_cell(sheet, "E34", "Priority", {}, left_bottom_border)

        # FFFFFFFFF
        add_cell(sheet, "F9", "Город", {}, left_bottom_border)
        add_cell(sheet, "F10", "City", {}, left_bottom_border)

        add_cell(sheet, "F19", "Город", {}, left_bottom_border)
        add_cell(sheet, "F20", "City", {}, left_bottom_border)

        sheet.merge_cells("F23:I23")
        add_cell(sheet, "F23", "Штемпель таможни", {}, left_bottom_border)

        sheet.merge_cells("F24:G24")
        add_cell(sheet, "F24", "Customs stamp", {}, left_bottom_border)

        sheet.merge_cells("F27:G27")
        add_cell(sheet, "F27", "Таможенный сбор", {}, left_bottom_border)

        sheet.merge_cells("F28:G28")
        add_cell(sheet, "F28", "Custom duty", {}, left_bottom_border)

        # GGGGGGGGG
        sheet.merge_cells("G9:I10")
        add_cell(sheet, "G9", "Garbsen, Niedersachsen", {bold: true}, center_left_wrapped_text)

        sheet.merge_cells("G19:I20")
        add_cell(sheet, "G19", "Almaty", {bold: true}, center_left_wrapped_text)

        # HHHHHHHHHHH
        sheet.merge_cells("H3:I3")
        add_cell(sheet, "H3", "Таможенная ссылка", {size: 6})

        sheet.merge_cells("H4:I4")
        add_cell(sheet, "H4", "отправителя (если имеется)", {size: 6}, top_wrapped_text_cell)

        sheet.merge_cells("H5:I5")
        add_cell(sheet, "H5", "Importer's reference ", {}, top_wrapped_text_cell)

        sheet.merge_cells("H6:I6")
        add_cell(sheet, "H6", "(tax code/VAT No./importer code) (optional)", {}, top_wrapped_text_cell)

        sheet.merge_cells("H48:I48")
        add_cell(sheet, "H48", "Дата и подпись", {}, left_bottom_border)

        sheet.merge_cells("H49:I49")
        add_cell(sheet, "H49", "отправителя", {}, left_bottom_border)

        sheet.merge_cells("H50:I50")
        add_cell(sheet, "H50", "Date and sender's ", {}, left_bottom_border)

        sheet.merge_cells("H51:I51")
        add_cell(sheet, "H50", "signature", {}, left_bottom_border)

        # JJJJJJJJJ

        sheet.merge_cells("J3:N3")
        add_cell(sheet, "J3", "№ посылки/посылок ", {}, left_bottom_border)

        sheet.merge_cells("J4:N4")
        add_cell(sheet, "J4", "(штриховой код, если имеется)", {}, left_bottom_border)

        sheet.merge_cells("J5:N5")
        add_cell(sheet, "J5", "No. of item", {}, left_bottom_border)

        sheet.merge_cells("J6:N6")
        add_cell(sheet, "J6", "(barcode, if any)", {}, left_bottom_border)

        sheet.merge_cells("J15:S15")
        add_cell(sheet, "J15", "Объявленная ценность - прописью", {}, left_bottom_border)

        sheet.merge_cells("J16:S16")
        add_cell(sheet, "J16", "Insured value - Words", {}, left_bottom_border)

        sheet.merge_cells("J17:S17")
        add_cell(sheet, "J17", "THIRTY EURO", {bold: true}, get_cell_style(sheet, nil, center_alignment, [:top, :right]))

        sheet.merge_cells("J18:S18")
        add_cell(sheet, "J18", "Сумма наложенного платежа - прописью", {}, left_bottom_border)

        sheet.merge_cells("J19:S19")
        add_cell(sheet, "J19", "Cash-on-delivery amount - words", {}, left_bottom_border)

        sheet.merge_cells("J21:S21")
        add_cell(sheet, "J21", "№ текущего почтового счета, центр чеков", {}, left_bottom_border)

        sheet.merge_cells("J22:S22")
        add_cell(sheet, "J22", "Giro account No. and Giro centre", {}, left_bottom_border)

        sheet.merge_cells("J24:S24")
        add_cell(sheet, "J24", "Реквизиты импортера (если имеются) (ИНН/№ НДС/индекс импортера) ", {}, left_bottom_border)

        sheet.merge_cells("J25:S25")
        add_cell(sheet, "J25", "(факультативно) ", {}, left_bottom_border)

        sheet.merge_cells("J27:V27")
        add_cell(sheet, "J27", "Importer's reference (if any) (taxcode/VAT No./importer code) (optional)", {}, left_bottom_border)

        sheet.merge_cells("J29:V29")
        add_cell(sheet, "J29", "№ телефона/факса/e-mail импортера (если известен)", {}, left_bottom_border)

        sheet.merge_cells("J30:V30")
        add_cell(sheet, "J30", "Importer's telephone/fax/e-mail (if known)", {}, left_bottom_border)

        sheet.merge_cells("J36:V36")
        add_cell(sheet, "J36", "Учреждение подачи/", {}, left_bottom_border)

        sheet.merge_cells("J36:V36")
        add_cell(sheet, "J36", "Учреждение подачи/", {}, left_bottom_border)

        sheet.merge_cells("J37:V37")
        add_cell(sheet, "J37", "Дата подачи", {}, left_bottom_border)

        sheet.merge_cells("J40:V40")
        add_cell(sheet, "J40", "Date of posting", {}, left_bottom_border)

        sheet.merge_cells("J41:M43")
        add_cell(sheet, "J41", "HAMBURG E KAZPOST", {bold: true}, get_cell_style(sheet, nil, center_alignment, [:top, :right]))

        sheet.merge_cells("J44:M46")
        add_cell(sheet, "J44", "28.08.2015", {bold: true}, get_cell_style(sheet, nil, center_alignment, [:top, :right]))

        sheet.merge_cells("J47:V47")
        add_cell(sheet, "J47", "Инструкции отправителя в случае невыдачи", {}, left_bottom_border)

        sheet.merge_cells("J48:V48")
        add_cell(sheet, "J48", "Sender's instruction in case of non-delivery", {}, left_bottom_border)




        # KKKKKKKKK
        sheet.merge_cells("K50:K51")
        add_cell(sheet, "K50", "", {bold: true, size: 16}, centered_cell)

        sheet.merge_cells("K53:K54")
        add_cell(sheet, "K53", "X", {bold: true, size: 16}, centered_cell)

        # LLLLLLLLL
        sheet.merge_cells("L7:V12")
        add_cell(sheet, "L7", "", {bold: true, size: 50}, centered_cell)

        sheet.merge_cells("L49:N49")
        add_cell(sheet, "L49", "Возвратить отправителю", {}, left_bottom_border)

        add_cell(sheet, "L50", "по истечении", {}, left_bottom_border)

        add_cell(sheet, "L51", "Return to the", {}, left_bottom_border)

        add_cell(sheet, "L52", "sender after", {}, left_bottom_border)

        sheet.merge_cells("L53:N53")
        add_cell(sheet, "L53", "Дослать получателю по", {}, left_bottom_border)

        sheet.merge_cells("L54:N54")
        add_cell(sheet, "L54", "нижеуказанному адресу", {}, left_bottom_border)

        sheet.merge_cells("L55:N55")
        add_cell(sheet, "L55", "Redirect to address", {}, left_bottom_border)

        sheet.merge_cells("L56:N56")
        add_cell(sheet, "L56", "below", {}, left_bottom_border)

        sheet.merge_cells("L57:N57")
        add_cell(sheet, "L57", "Адрес/Address", {}, left_bottom_border)

        sheet.merge_cells("L60:V60")
        add_cell(sheet, "L60", "KZ Post", {}, left_bottom_border)

        # MMMMMMMMM
        sheet.merge_cells("M50:M51")
        add_cell(sheet, "M50", "", {bold: true, size: 16}, centered_cell)

        # NNNNNNNNN
        add_cell(sheet, "M50", "Дней", {}, left_bottom_border)
        add_cell(sheet, "M50", "Days", {}, left_bottom_border)

        sheet.merge_cells("N36:R36")
        add_cell(sheet, "N36", "Количество посылок", {}, left_bottom_border)

        sheet.merge_cells("N37:R37")
        add_cell(sheet, "N37", "Number of parcels", {}, left_bottom_border)

        sheet.merge_cells("N40:V40")
        add_cell(sheet, "N40", "Объявленная ценность в СПЗ", {}, left_bottom_border)

        sheet.merge_cells("N41:V41")
        add_cell(sheet, "N41", "insured value SDR", {}, left_bottom_border)

        sheet.merge_cells("N43:T43")
        add_cell(sheet, "N43", "Общий вес брутто посылки/посылок", {}, left_bottom_border)

        sheet.merge_cells("N44:T46")
        add_cell(sheet, "N44", "Total gross weight of the parcel(s)", {}, left_bottom_border)


        # OOOOOOOOO
        sheet.merge_cells("O56:O57")
        add_cell(sheet, "O56", "X", {bold: true, size: 16}, centered_cell)

        # PPPPPPPPP
        sheet.merge_cells("P56:S56")
        add_cell(sheet, "P56", "Наземным путем/S.A.L.", {}, left_bottom_border)

        sheet.merge_cells("P57:S57")
        add_cell(sheet, "P57", "by surface/S.A.L", {}, left_bottom_border)

        # QQQQQQQQQ
        sheet.merge_cells("Q50:Q51")
        add_cell(sheet, "Q50", "", {bold: true, size: 16}, centered_cell)

        sheet.merge_cells("Q53:Q54")
        add_cell(sheet, "Q53", "", {bold: true, size: 16}, centered_cell)


        # RRRRRRRRR
        sheet.merge_cells("R50:V50")
        add_cell(sheet, "R50", "Возвратить сразу же отправителю", {}, left_bottom_border)

        sheet.merge_cells("R51:V51")
        add_cell(sheet, "R51", "Возвратить сразу же отправителю", {}, left_bottom_border)

        sheet.merge_cells("R51:V51")
        add_cell(sheet, "R51", "Возвратить сразу же отправителю", {}, left_bottom_border)

        sheet.merge_cells("R53:V53")
        add_cell(sheet, "R53", "Обрабатывать как отказное", {}, left_bottom_border)

        sheet.merge_cells("R54:V54")
        add_cell(sheet, "R53", "Treat as abondoned", {}, left_bottom_border)

        # SSSSSSSSS

        sheet.merge_cells("S36:V36")
        add_cell(sheet, "S36", "Сертификатов и счетов", {}, left_bottom_border)

        sheet.merge_cells("S37:V37")
        add_cell(sheet, "S37", "certificates and invoices", {}, left_bottom_border)

        # TTTTTTTTT
        add_cell(sheet, "T15", "Цифрами", {}, left_bottom_border)
        add_cell(sheet, "T16", "figures", {}, left_bottom_border)

        sheet.merge_cells("T17:V17")
        add_cell(sheet, "T17", "30.00", {bold: true}, get_cell_style(sheet, nil, center_alignment, [:top, :right]))

        add_cell(sheet, "T18", "Цифрами", {}, left_bottom_border)
        add_cell(sheet, "T19", "figures", {}, left_bottom_border)

        # UUUUUUUUU
        add_cell(sheet, "U43", "Тарифы", {}, left_bottom_border)
        add_cell(sheet, "U44", "Charges", {}, left_bottom_border)

        sheet.merge_cells("U56:Q57")
        add_cell(sheet, "U56", "", {bold: true, size: 16}, centered_cell)

        # VVVVVVVVV
        add_cell(sheet, "V56", "Авиа", {}, left_bottom_border)
        add_cell(sheet, "V57", "by air", {}, left_bottom_border)

      end
    end
    p.serialize result_path
  end
end

form = FormCP71.new
form.process
form.open_file