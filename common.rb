module Common

  def add_logo(sheet, start_at_row, start_at_col, file_path = nil)
    logo = File.expand_path(file_path || 'images/kaz_logo.png')
    sheet.add_image(image_src: logo, noSelect: false, noMove: false) do |image|
      image.height = 22
      image.width = 130
      image.start_at start_at_row, start_at_col
    end
  end

  def add_signature(sheet, start_at_row, start_at_col, file_path = nil)
    signature = File.expand_path(file_path || 'images/signature.png')
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
    # if value.is_a?(Axlsx::RichText)
    #   cell.type = :richtext
    #   value.cell = cell
    # end
    cell
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

end