import json
import pymysql



def read_basic_info():
    # {hostname:{if_index:[if_name,if_description]}}
    json_data = json.loads(open('basic_info.json').read())
    basic_dict = {}
    for js in json_data:
        data = js.get('data')[0]
        basic_dict.setdefault(data['hostname'], {})
        for index in data['if_desc']:
            if index in data['if_index']:
                basic_dict[data['hostname']].update({index: [data['if_index'][index], data['if_desc'][index]]})
    return basic_dict


def write_in_csv(data):
    csv = open('data.csv', "w")
    title = "hostname, if_index,if_name,if_description\n"
    csv.write(title)
    for k, v in data.items():
        for index, desc_name in v.items():
            row = k + "," + index + "," + desc_name[0] + "," + desc_name[1] + "\n"
            csv.write(row)


def save_to_database(filename):
    db = pymysql.connect('localhost', 'root', 'root', 'test')
    cursor = db.cursor()
    sql = 'LOAD DATA INFILE "%s" INTO TABLE basic_info FIELDS TERMINATED BY "," IGNORE 1 LINES' % filename
    print(sql)
    cursor.execute(sql)
    db.commit()
    db.close()


if __name__ == '__main__':
    basic_data = read_basic_info()
    write_in_csv(basic_data)
    save_to_database(r'C:/Users/xinzha4/PycharmProjects/test_json_data/data.csv')
    print('end****************')
    pass
