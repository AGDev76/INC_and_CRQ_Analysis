from openpyxl import load_workbook, workbook
from openpyxl.formatting.rule import ColorScaleRule
from fuzzywuzzy import fuzz, process
import sys
import getopt


def main(argv):
    inputfile = ''
    outputfile = ''
    try:
        opts, args = getopt.getopt(argv, "hi:o:", ["ifile=", "ofile="])
    except getopt.GetoptError:
        print('Usage: python3 yourfile.py -i <inputfile> -o <outputfile>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print('Usage: python3 yourfile.py -i <inputfile> -o <outputfile>')
            sys.exit()
        elif opt in ("-i", "--ifile"):
            inputfile = arg
        elif opt in ("-o", "--ofile"):
            outputfile = arg

    if inputfile == '':
        print("Usage: python3 yourfile.py -i <inputfile> -o <outputfile>")
        return
    else:
        workbook = load_workbook(filename=inputfile)
        sheet = workbook.active
        server_names = []
        res_counter = 2
        det_counter = 2

        try:
            with open("keywords.txt", "r") as keywords:
                for row in keywords:
                    data = row.strip("\n")
                    server_names.append(data)
        except Exception as e:
            print(e)

        if sheet["R1"].value == None and sheet["S1"].value == None and sheet["T1"].value == None and sheet["U1"].value == None:
            sheet["R1"] = "Res_Server_Name"
            sheet["S1"] = "Res_Word_Match"
            sheet["T1"] = "Det_Server_Name"
            sheet["U1"] = "Det_Word_Match"
            try:
                workbook.save(filename=outputfile)
            except Exception as e:
                print(e)

        for value in sheet.iter_rows(min_row=2, min_col=15, max_col=15, values_only=True):
            if value[0] != None:
                result = process.extractOne(
                    value[0], server_names, scorer=fuzz.partial_ratio)
                percent = result[1]
                if percent > 80 and sheet["R" + str(res_counter)] != None and sheet["S" + str(res_counter)] != result[1] != None:
                    sheet["R" + str(res_counter)] = result[0]
                    sheet["S" + str(res_counter)] = result[1]
                    try:
                        workbook.save(filename=outputfile)
                    except Exception as e:
                        print(e)
                else:
                    sheet["R" + str(res_counter)] = " "
                    workbook.save(filename=outputfile)
                res_counter += 1
            else:
                res_counter += 1

        for value in sheet.iter_rows(min_row=2, min_col=16, max_col=16, values_only=True):
            if value[0] != None:
                result = process.extractOne(
                    value[0], server_names, scorer=fuzz.partial_ratio)
                percent = result[1]
                if percent > 80 and sheet["T" + str(det_counter)] != None and sheet["U" + str(det_counter)] != result[1] != None:
                    sheet["T" + str(det_counter)] = result[0]
                    sheet["U" + str(det_counter)] = result[1]
                    try:
                        workbook.save(filename=outputfile)
                    except Exception as e:
                        print(e)
                det_counter += 1
            else:
                det_counter += 1

        color_scale_rule = ColorScaleRule(start_type="num",
                                          start_value=80,
                                          start_color="00FF0000",
                                          mid_type="num",
                                          mid_value=90,
                                          mid_color="00FFFF00",
                                          end_type="num",
                                          end_value=100,
                                          end_color="0000FF00")

        sheet.conditional_formatting.add("S2:S25000", color_scale_rule)
        sheet.conditional_formatting.add("U2:U25000", color_scale_rule)
        workbook.save(filename=outputfile)


if __name__ == "__main__":
    main(sys.argv[1:])
