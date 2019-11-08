import sys
import xlrd
import json
from datetime import datetime


def main():
    if len(sys.argv) < 2:
        print('error: filename required!')
        return 1

    try:
        workbook = xlrd.open_workbook(sys.argv[1])
    except:
        print('error: file not found!')    

    data = []

    for sheet in workbook.sheets():

        for row_index in range(sheet.nrows):
            cells = sheet.row_values(row_index)
            data.append(cells)

    raw_cities = [row[1] for row in data]
    unordered_cities = set()
    ordered_cities = [x for x in raw_cities if not (x in unordered_cities or unordered_cities.add(x))]

    cities = []

    for c in ordered_cities:
        city = {'n': c}
        city['areas'] = []
        cities.append(city)
        matched_city = [x for x in data if x[1] == c]
        raw_areas = [x[2] for x in matched_city]
        unordered_areas = set()
        ordered_areas = [x for x in raw_areas if not (x in unordered_areas or unordered_areas.add(x))]

        for a in ordered_areas:
            area = {'n': a}
            area['roads'] = []
            city['areas'].append(area)
            matched_area = [x for x in matched_city if (x[1] == c and x[2] == a)]
            raw_roads = [x[3] for x in matched_area]
            unordered_roads = set()
            ordered_roads = [x for x in raw_roads if not (x in unordered_roads or unordered_roads.add(x))]

            for r in ordered_roads:
                road = {'n': r}
                road['zs'] = []
                area['roads'].append(road)
                matched_road = [x for x in matched_area if (x[1] == c and x[2] == a and x[3] == r)]

                for m in matched_road:
                    zipcode = {'z': m[0], 'd':m[4]}
                    road['zs'].append(zipcode)

    now = datetime.now()
    result = {'cities': cities, 'update': now.strftime("%Y-%m-%d %H:%M:%S")}

    with open('tw-zipcode.json', 'w') as outfile:
        json.dump(result, outfile, ensure_ascii=False)
        
    pass


if __name__ == '__main__':
    main()
