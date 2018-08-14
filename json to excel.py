import json
import xlwt


def get_columns_data():
    columns = json.load(open('columns.json', 'r', encoding='UTF-8'))
    excel_name = columns.get('excel_name')
    sheet_name = columns.get('sheet_name')
    _titles = columns.get('titles')
    return _titles, excel_name, sheet_name


def get_json_data():
    data = json.load(open('data.json', 'rb'))
    return data


def build_columns_data(title):
    data = get_json_data()
    mid_value = []
    for k, v in title.items():
        for datum in data:
            if datum.get(v.get('key')) == '':
                value = v.get('default')
            else:
                value = datum.get(v.get('key'))
            mid_value.append(value)
    res = []
    j = 1
    num = int(len(mid_value) / len(title))
    for i in range(int(len(title))):
        res.append(mid_value[i * num:num * j:1])
        i += 1
        j += 1
    count = 0
    for k, v in title.items():
        del v['key']
        del v['default']
        title[k]['value'] = res[count]
        count += 1
    # {'J': {'label': 'IP交换机位置', 'value': []}
    return title


def write_excel(value, excel_name, sheet_name):
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    ws = book.add_sheet(sheet_name, cell_overwrite_ok=True)
    for k in value:
        ws.write(0, ord(k) - 65, value[k]['label'], set_style('宋体', 255, 64, 2, 0x37, False))
        length = []
        for i in range(len(value[k]['value'])):
            length.append(len(value[k]['value'][i].encode('utf-8')) + 2)
            ws.write(i + 1, ord(k) - 65, value[k]['value'][i], set_style('宋体', 220, 64, 2, 1, False))
        ws.col(ord(k) - 65).width = (max(length) + 5) * 256
    book.save(excel_name)


def set_style(name, height, color, alignment, i, bold=False):
    style = xlwt.XFStyle()  # 初始化样式
    font = xlwt.Font()  # 为样式创建字体
    # 字体类型：比如宋体、仿宋也可以是汉仪瘦金书繁
    font.name = name
    # 是否为粗体
    font.bold = bold
    # 设置字体颜色
    font.colour_index = color
    # 字体大小
    font.height = height
    # 字体是否斜体
    font.italic = False
    # # 字体下划,当值为11时。填充颜色就是蓝色
    # font.underline = 0
    # 字体中是否有横线struck_out
    font.struck_out = False
    # 定义格式
    style.font = font
    # 单元格背景色
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = i
    style.pattern = pattern
    # 居中
    style.alignment.horz = alignment
    # alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平方向
    # alignment.vert = xlwt.Alignment.VERT_TOP  # 垂直方向
    # 设置边框
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    style.borders = borders
    return style


if __name__ == '__main__':
    titles, excel, sheet = get_columns_data()
    val = build_columns_data(titles)
    write_excel(val, excel, sheet)



