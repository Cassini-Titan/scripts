import subprocess
import typing
import os
import shutil
from pathlib import Path


def redStr(str: str):
    return " \033[1;31m " + str + " \033[0m "


def greenStr(str: str):
    return " \033[1;32m " + str + " \033[0m "


def yelloStr(str: str):
    return " \033[1;33m " + str + " \033[0m "


def buleStr(str: str):
    return " \033[1;34m " + str + " \033[0m "


def errorOut(str: str):
    print(redStr("[Error]: " + str))


def warningOut(str: str):
    print(yelloStr("[Warning]: " + str))


def successOut(str: str):
    print(greenStr("[Success]: " + str))


def infOut(str: str):
    print(buleStr("[Info]: " + str))


def runCmd(command: str) -> typing.Tuple:
    subp = subprocess.Popen(
        command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
    )
    res, err = subp.communicate()
    if subp.poll() == 0:
        successOut(command)
        return 0, res.decode("utf-8")
    else:
        errorOut(err.decode("utf-8"))
        return -1, err.decode("utf-8")


if __name__ == "__main__":
    cwd = os.getcwd()
    infOut("Current directory is " + cwd)
    sub_folder_list = [f.path for f in os.scandir(cwd) if f.is_dir()]
    for sub_folder in sub_folder_list:
        sub_folder_name = sub_folder.split("/")[-1]
        new_folder = cwd + "/py-" + sub_folder_name
        if new_folder not in sub_folder_list and sub_folder_name.find("py-") == -1:
            runCmd(f"cp -r {sub_folder} {new_folder}")
            try:
                shutil.rmtree(new_folder + "/.git")
                checkpoint_dir_list = list(Path(new_folder).rglob(".ipynb_checkpoints"))
                for checkpoint_dir in checkpoint_dir_list:
                    shutil.rmtree(checkpoint_dir)
            finally:
                ipynb_list = list(Path(new_folder).rglob("*.ipynb"))
                for ipynb in ipynb_list:
                    ipynb = str(ipynb)
                    runCmd(f"jupyter nbconvert --to script {ipynb}")
                    os.remove(ipynb)
