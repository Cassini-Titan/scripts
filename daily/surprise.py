import argparse
import sys
import csv
import typing
import time
import random
from pathlib import Path
import pandas as pd


class Person:
    def __init__(self, row, header) -> None:
        self.__dict__ = dict(zip(header, row))

    def get_person(self) -> dict:
        return self.__dict__

    def get_name(self) -> str:
        return self.__dict__["Name"]

    def get_count(self) -> str:
        return self.__dict__["Count"]

    def get_last_time(self) -> str:
        return self.__dict__["LastTime"]

    def plus(self) -> None:
        count = int(self.get_count())
        self.__dict__["Count"] = str(count + 1)
        self.__dict__["LastTime"] = time.time()
        print("Plus cnt for {}".format(self.get_name()))

    def minus(self) -> None:
        count = int(self.get_count())
        self.__dict__["Count"] = str(count - 1)
        print("Minus cnt for {}".format(self.get_name()))

    def __repr__(self) -> str:
        return self.get_name()


def init_cnt(excel_file: str) -> None:
    df = pd.read_excel(excel_file, sheet_name="Sheet1", index_col=None, usecols="A:C")
    df.to_csv("cnt.csv", encoding="utf-8", index=False)


def read_cnt(csv_file: str) -> typing.List:
    data = list(csv.reader(open(csv_file, "r", encoding="utf-8")))
    people = [Person(row, data[0]) for row in data[1:]]
    # for person in people:
    #     if person.Count == "2" or person.Count == "1":
    #         person.LastTime = time.time()
    #     else:
    #         person.LastTime = 0
    return people


def write_cnt(csv_file: str, people: list) -> None:
    keys = ["Name", "Count", "LastTime"]
    with open(csv_file, "w", encoding="utf-8", newline="") as f:
        f_csv = csv.DictWriter(f, keys)
        f_csv.writeheader()
        for person in people:
            f_csv.writerow(person.get_person())


def read_record_from_excel(excel_file: str, column: int) -> typing.List:
    if column == None:
        column = 1
    col = chr(column + 64).upper()
    df = pd.read_excel(excel_file, usecols=col, sheet_name="Sheet1")
    return df.iloc[:, 0].tolist()


def plus_cnt(record_file: str, cnt_file: str, column: int = 1) -> None:
    people = read_cnt(cnt_file)
    records = read_record_from_excel(record_file, column)
    for person in people:
        name = person.get_name()
        if name in records:
            person.plus()
    write_cnt(cnt_file, people)


def minus_cnt(record_file: str, cnt_file: str, column: str) -> None:
    people = read_cnt(cnt_file)
    records = read_record_from_excel(record_file, column)
    for person in people:
        name = person.get_name()
        if name in records:
            person.minus()
    write_cnt(cnt_file, people)


def get_same_count_number(people: list, base_index: int) -> int:
    base_count = people[base_index].get_count()
    cnt = base_index
    while cnt < len(people):
        if people[cnt].get_count() == base_count:
            cnt += 1
        else:
            break
    return cnt - base_index


def sort(csv_file: str) -> typing.List:
    people = read_cnt(csv_file)
    sorted_by_count = sorted(people, key=lambda x: x.get_count())
    sorted_people = []
    fragments = []
    cnt = 0

    # sort by count
    while cnt < len(sorted_by_count):
        base_number = get_same_count_number(sorted_by_count, cnt)
        fragments.append(base_number)
        cnt += base_number

    # sort by last time
    base_index = 0
    for fragment in fragments:
        sublist = sorted_by_count[base_index : base_index + fragment]
        sublist = sorted(sublist, key=lambda x: x.get_last_time())
        # shuffle when count and last time are same
        sub_index = 0
        while sub_index < fragment - 1:
            next_index = sub_index + 1
            if (
                sublist[sub_index].get_last_time()
                != sublist[next_index].get_last_time()
            ):
                sub_index += 1
            else:
                tmp = sublist[sub_index:]
                random.shuffle(tmp)
                sublist[sub_index:] = tmp
                break
        for person in sublist:
            sorted_people.append(person)
        base_index += fragment

    return sorted_people


def surprise(require_number: int, csv_file: str) -> typing.List:
    people = sort(csv_file)
    return people[:require_number]


if __name__ == "__main__":
    csv_file = Path.absolute(Path(__file__)).parent.parent.parent / "cnt.csv"
    record_file = Path.absolute(Path(__file__)).parent.parent.parent / "record.xlsx"
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-p", "--plus", type=int, help="plus which column", nargs="?"
    )
    parser.add_argument("-m", "--minus", type=int, help="minus which column", nargs="?")
    parser.add_argument("-n", "--number", type=int, help="required number", nargs="?")
    args, _ = parser.parse_known_args()
    plus_col = args.plus
    minus_col = args.minus
    required_number = args.number
    if ("-p" in sys.argv):
        plus_cnt(record_file, csv_file, plus_col)
    if minus_col is not None:
        minus_cnt(record_file, csv_file, minus_col)
    if ("-n" in sys.argv):
        if required_number is not None:
            print(required_number)
            print(surprise(required_number, csv_file))
        else:
            print("Please input require number")