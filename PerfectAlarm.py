# -*- coding: utf-8 -*-
'''
* @author UUGTech
* @copyright 2020 by UUGTech
'''
import choosedocx
import french_fries_paradise
import make_schedule
import sys


if __name__ == "__main__":
    args = sys.argv
    time = str(args[1])
    if len(args)==2 and len(time)==5 and time[2] == ":" and int(time[0:2]) < 24 and int(time[3:5])<=59:
        fname = choosedocx.choose()
        make_schedule.make_alarm_task(str(args[1]), fname, french_fries_paradise.echo_path())
    else:
        print("python PerfectAlarm.py 05:30   の表記で時刻を指定してください")