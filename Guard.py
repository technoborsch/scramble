import tempfile
import time
import os
import base64
import random

import config


temp_dir = os.path.join(tempfile.gettempdir(), "ce204fRet")
temp_file = os.path.join(temp_dir, "config_path.cfg")


def encode(target_time: float, program_version: float):
    time_text = str(int(target_time))
    version_text = ""
    for char in str(program_version):
        version_text += str(ord(char) - int(config.CHAR_OFFSET))
    result = ""
    for i in range(len(time_text)):
        if i < len(version_text):
            result += time_text[i] + version_text[i]
        else:
            result += time_text[i] + str(random.randrange(0, 9))
    result = "8" + result[::-1]
    return base64.b64encode(str.encode(result)).decode("utf-8")


def decode(string):
    cipher = base64.b64decode(string).decode("utf-8")
    inter = cipher[1:][::-1]
    target_time = ""
    program_version = ""
    for i in range(len(inter)):
        if i % 2:
            program_version += inter[i]
        else:
            target_time += inter[i]
    program_version = program_version[:6]
    version = ""
    version_list = [program_version[j:j+2] for j in range(0, len(program_version), 2)]
    for char in version_list:
        version += chr(int(char) + int(config.CHAR_OFFSET))
    return int(target_time), float(version)


def set_temp_value(value):
    if not os.path.exists(temp_dir):
        os.mkdir(temp_dir)
    with open(temp_file, "w") as f:
        f.write(value)


def get_temp_value():
    if os.path.exists(temp_file):
        with open(temp_file, "r") as f:
            return f.read()


def check():
    now_time = int(time.time())  # узнаем текущее время
    result = False
    try:
        this_target_time, current_version = decode(config.SECRET)  # узнаем текущую версию и время просрочки
    except Exception:
        return False
    temp_value = get_temp_value()  # проверяем, есть ли временный файл
    if this_target_time > now_time:  # если текущее время меньше времени просрочки, идем дальше
        if temp_value:  # если есть, анализируем
            try:
                last_launch_time, last_version = decode(temp_value)
            except Exception:
                return False
            if last_launch_time < now_time:  # если последний запуск был раньше текущего времени (обязательное условие!)
                if current_version >= last_version:  # если текущая версия не меньше использованной ранее
                    result = True
        else:  # если нет, то ставим и разрешаем запуск (слабое место)
            set_temp_value(encode(now_time, config.PROGRAM_VERSION))
            result = True
    else:  # запускают после просрочки
        if temp_value:  # значит, запускали до этого
            try:
                last_launch_time, last_version = decode(temp_value)
            except Exception:
                return False
            if last_launch_time < now_time:  # ставим метку только если текущее время запуска еще больше предыдущего
                set_temp_value(encode(now_time, config.PROGRAM_VERSION))
    return result


if __name__ == "__main__":
    print(config.SECRET)
