import csv

# ----------------------------------------------------------------------


def csv_dict_reader(file_obj):

    """
    Read a CSV file using csv.DictReader
    """
    reader = csv.DictReader(file_obj, delimiter=',')
    for line in reader:
        if line["Task"] == "03.1-On - Development":
            print "INVALID DATA FOUND"
        else:
            print "CORRECT DATA FOUND"

        print(line["Employee"]),
        print(line["Task"])


# ----------------------------------------------------------------------
if __name__ == "__main__":
    with open("C:\\PycharmProjects\\mytime_emp_rh_201806_all_projects.csv") as f_obj:
        csv_dict_reader(f_obj)




