import argparse
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

    def update(self) -> None:
        count = int(self.get_count())
        self.__dict__["Count"] = str(count + 1)
        self.__dict__["LastTime"] = time.time()

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


def read_record_from_excel(excel_file: str) -> list:
    df = pd.read_excel(excel_file, usecols="A", sheet_name="Sheet1")
    return df["Record"].tolist()


def update_cnt(record_file: str, cnt_file: str) -> None:
    people = read_cnt(cnt_file)
    records = read_record_from_excel(record_file)
    for person in people:
        name = person.get_name()
        if name in records:
            person.update()
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


def sort(csv_file: str) -> list:
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
    csv_file = Path.absolute(Path(__file__)).parent.parent / "cnt.csv"
    print(csv_file)
    parser = argparse.ArgumentParser()
    parser.add_argument("-r", "--record", help="record file path", nargs="?")
    parser.add_argument("-n", "--number", type=int, help="required number")
    parser.parse_args()
    record_file = parser.parse_args().record
    required_number = parser.parse_args().number
    if record_file:
        update_cnt(record_file, csv_file)
    if required_number:
        people = surprise(required_number, csv_file)
        print(people)
    else:
        print("Please input require number")
