from json import load, dumps
from os import path
from openpyxl import load_workbook


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
        "event_data": open_workbook_as_dictionaries("./reference/event_survey_data.xlsx"),
    }

    with open(filepath, 'w') as file:
        file.write(dumps(data))

    return data


def get_survey_response(pid, event_data):
    for d in event_data:
        if d.get("PID") == pid:
            return d


def build_row(d, event_data, headers, size):
    row = [d.get(key) for key in headers.get("participant_info")]
    pid = d.get("PID")
    event_d = get_survey_response(pid, event_data)
    if event_d:
        for key in headers.get("survey_headers"):
            row.append(event_d.get(key))
    else:
        while len(row) < size:
            row.append("")
    return row


def stitch_files(data):
    with open("./formatting.json") as file:
        headers = load(file)

    tabular_data = [[header for header in headers.get(
        "participant_info")] + [header for header in headers.get("survey_headers")]]
    for d in data.get('cm_data'):
        row = build_row(d, data.get("event_data"),
                        headers, len(tabular_data[0]))
        tabular_data.append(row)

    return tabular_data


def load_settings(filepath):
    with open(filepath) as file:
        return load(file)


def get_tabular_data(**kwargs):
    b = kwargs.get("overwrite", False)
    data = load_file_info(overwrite=b)
    tabular_data = stitch_files(data)
    return tabular_data


def main():
    data = load_file_info(overwrite=True)
    tabular_data = stitch_files(data)


if __name__ == "__main__":
    main()
