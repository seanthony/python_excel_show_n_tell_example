from json import load, dumps
from os import path
from openpyxl import load_workbook


def load_settings(filepath):
    with open(filepath) as file:
        return load(file)


def open_workbook_as_dictionaries(filepath):
    d_map = {
        'Strongly Agree': 7,
        'Agree': 6,
        'Somewhat Agree': 5,
        'Neutral': 4,
        'Somewhat Disagree': 3,
        'Disagree': 2,
        'Strongly Disagree': 1
    }

    wb = load_workbook(filepath)
    sheet = wb.active

    l = [[cell.value for cell in row] for row in sheet]
    header, values = l[0], l[1:]

    l_d = []
    for row in values:
        d = {}
        for i in range(len(header)):
            key = header[i]
            value = row[i]
            if value in d_map.keys():
                value = d_map.get(value)
            d[key] = value
        l_d.append(d)

    return l_d


def load_file_info(**kwargs):
    overwrite = kwargs.get('overwrite', None)

    filepath = './cmInfo.json'
    if path.isfile(filepath) and not overwrite:
        with open(filepath) as file:
            return load(file)

    data = {
        "cm_data": open_workbook_as_dictionaries("./reference/cms.xlsx"),
        "event_data": open_workbook_as_dictionaries("./reference/cms.xlsx"),
    }

    with open(filepath, 'w') as file:
        file.write(dumps(data))

    return data


def main():
    settings = load_settings("./settings.json")
    data = load_file_info(overwrite=False)
    print(data)


if __name__ == "__main__":
    main()
