from XlsxWriter import xlsxwriter


def get_data(input_file):
    f = open(input_file)
    line_list = []
    lines = []

    for line in f.readlines():
        for item in line.split(","):
            line_list.append(item)
        if line[-2] == ',':
            lines.append(line_list)
            line_list = []
        else:
            line_list.append("same_item")

    for line in lines:
        for index, item in enumerate(line):
            if item == "same_item":
                line[index - 1] += line[index + 1]
                del line[index + 1]
                del line[index]
    return lines


workbook = xlsxwriter.Workbook("result.xlsx")
worksheet = workbook.add_worksheet()
title_format = workbook.add_format({'bold': True})
title_format.set_bottom(5)
title_format.set_top(5)
title_format.set_left(2)
title_format.set_right(2)
guardian_format = title_format
guardian_format.set_left(5)
guardian_format.set_right(5)

# initialize fist row
column_list = ["Title", "Guest Last Name", "Guest First Name", "Photo", "Room", "VIP", "Arrival", "Departure",
               "Email /Club"]
for i, col in enumerate(column_list):
    worksheet.write(0, i, col, title_format)
worksheet.write(0, 9, "Guardian Angel", guardian_format)

data = get_data("input.csv")
col = 0
row = 1
title_format = workbook.add_format()
title_format.set_bottom(5)
title_format.set_top(5)
title_format.set_left(2)
title_format.set_right(2)
title_format.set_align("left")
title_format.set_align("top")
title_format.set_text_wrap()
name_format = workbook.add_format()
name_format.set_right(2)
name_format.set_left(2)
name_format.set_top(5)
name_format.set_bottom(3)
name_format.set_align("left")
name_format.set_align("top")
name_format.set_text_wrap()
comment_format = workbook.add_format()
comment_format.set_right(2)
comment_format.set_left(2)
comment_format.set_top(3)
comment_format.set_bottom(5)
comment_format.set_align("left")
comment_format.set_align("top")
general_format = workbook.add_format()
general_format.set_right(2)
general_format.set_left(2)
general_format.set_top(5)
general_format.set_bottom(5)
general_format.set_align("left")
general_format.set_align("top")
guardian_format.set_bold(False)
guardian_format.set_align("left")
guardian_format.set_align("top")

room_dictionary = {}
for person in data:
    if person[4] in room_dictionary:
        room_dictionary[person[4]][0] += '\n' + person[0]
        room_dictionary[person[4]][1] += '\n' + person[1].replace(chr(34), "")
        room_dictionary[person[4]][2] += '\n' + person[2].replace(chr(34), "")
        new_content = room_dictionary[person[4]]
        temp_row = int(new_content[-1])
        # Title
        worksheet.write(temp_row, 0, new_content[0], title_format)
        # Guest Last Name
        worksheet.write(temp_row, 1, new_content[1], name_format)
        # Guest First Name
        worksheet.write(temp_row, 2, new_content[2], name_format)
        # Comment
        # Photo
        # Room
        # VIP
        # Arrival
        # Departure
        # Email / Club
        # Guardian Angel
    elif person[1]:
        # Title
        worksheet.merge_range(row, 0, row+2, 0, '', title_format)
        worksheet.write(row, 0, person[0], title_format)
        # Guest Last Name
        person[1] = person[1].replace(chr(34), "")
        worksheet.merge_range(row, 1, row + 1, 1, '', name_format)
        worksheet.write(row, 1, person[1], name_format)
        # Guest First Name
        person[2] = person[2].replace(chr(34), "")
        worksheet.merge_range(row, 2, row + 1, 2, '', name_format)
        worksheet.write(row, 2, person[2], name_format)
        # Comment
        worksheet.merge_range(row + 2, 1, row + 2, 2, '', comment_format)
        # Photo
        worksheet.merge_range(row, 3, row + 2, 3, '', general_format)
        # Room
        worksheet.merge_range(row, 4, row + 2, 4, '', general_format)
        worksheet.write(row, 4, person[4], general_format)
        # VIP
        worksheet.merge_range(row, 5, row + 2, 5, '', general_format)
        worksheet.write(row, 5, person[5], general_format)
        # Arrival
        worksheet.merge_range(row, 6, row + 2, 6, '', general_format)
        worksheet.write(row, 6, person[6], general_format)
        # Departure
        worksheet.merge_range(row, 7, row + 2, 7, '', general_format)
        worksheet.write(row, 7, person[7], general_format)
        # Email / Club
        worksheet.merge_range(row, 8, row + 2, 8, '', general_format)
        if person[7]:
            worksheet.write(row, 8, "Y", general_format)
        else:
            worksheet.write(row, 8, "N", general_format)
        # Guardian Angel
        worksheet.merge_range(row, 9, row + 2, 9, '', guardian_format)
        worksheet.write(row, 9, person[9], guardian_format)
        person.append(str(row))
        row += 3
        room_dictionary[person[4]] = person


workbook.close()
