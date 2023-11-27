# for row_num, row in enumerate(template_ws.iter_rows()):
#                 for col_num, cell in enumerate(row):
#                     dest_cell =dest_ws.cell(row=row_num+1, column=col_num+1)
#                     dest_cell.value = cell.value
#                     dest_cell.number_format = cell.number_format
#                     dest_cell.font = openpyxl.styles.Font(
#                         name=cell.font.name,
#                         size=cell.font.size,
#                         bold=cell.font.bold,
#                         italic=cell.font.italic,
#                         underline=cell.font.underline,
#                         strike=cell.font.strike,
#                         color=cell.font.color
#                     )
#                     dest_cell.alignment = openpyxl.styles.Alignment(
#                         horizontal=cell.alignment.horizontal,
#                         vertical=cell.alignment.vertical,
#                         text_rotation=cell.alignment.textRotation,
#                         wrap_text=cell.alignment.wrapText,
#                         shrink_to_fit=cell.alignment.shrinkToFit,
#                         indent=cell.alignment.indent,
#                         relativeIndent=cell.alignment.relativeIndent,
#                         justifyLastLine=cell.alignment.justifyLastLine,
#                         readingOrder=cell.alignment.readingOrder,
#                     )
#                     dest_cell.border = openpyxl.styles.Border(
#                         left=cell.border.left,
#                         right=cell.border.right,
#                         top=cell.border.top,
#                         bottom=cell.border.bottom,
#                         diagonal=cell.border.diagonal,
#                         diagonal_direction=cell.border.diagonal_direction,
#                         start=cell.border.start,
#                         end=cell.border.end
#                     )
#                     dest_cell.fill = openpyxl.styles.PatternFill(
#                         fill_type=cell.fill.fill_type,
#                         start_color=cell.fill.start_color,
#                         end_color=cell.fill.end_color
#                     )
