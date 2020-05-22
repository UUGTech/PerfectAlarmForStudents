# -*- coding: utf-8 -*-
'''
* @author UUGTech
* @copyright 2020 by UUGTech
'''
import subprocess
import os
import sys
import xml.etree.ElementTree as ET
import datetime

def make_alarm_task(time, fname, path):
    # open xml
    tree = ET.parse("./template.xml")
    root = tree.getroot()
    # set date data
    now = datetime.datetime.now()
    date = datetime.datetime(now.year, now.month, now.day, int(time[0:2]), int(time[3:5]), 00)
    if date < now :
        date = date + datetime.timedelta(days=1)
    # edit template.xml
    for start_boundary in root.iter("{http://schemas.microsoft.com/windows/2004/02/mit/task}StartBoundary"):
        start_boundary.text = "{:0=4d}-{:0=2d}-{:0=2d}T{:0=2d}:{:0=2d}:00".format(date.year, date.month, date.day, date.hour, date.minute)
    for d in root.iter("{http://schemas.microsoft.com/windows/2004/02/mit/task}Date"):
        d.text = "{:0=4d}-{:0=2d}-{:0=2d}T{:0=2d}:{:0=2d}:00".format(now.year, now.month, now.day, now.hour, now.minute)
    for command in root.iter("{http://schemas.microsoft.com/windows/2004/02/mit/task}Command"):
        python_path = sys.executable
        command.text = python_path
    for arguments in root.iter("{http://schemas.microsoft.com/windows/2004/02/mit/task}Arguments"):
        script_path = path
        arguments.text = script_path + " " + fname
    # wriete xml file as "Alarm.xml"
    tree.write("./Alarm.xml", encoding="UTF-16", xml_declaration=True)
    # call schtasks.exe to make task
    xmlFile = os.path.dirname(os.path.abspath(__file__)) + "/Alarm.xml"
    username = "RUNUSER"
    process = "schtasks /create /tn \"Alarm\" /xml " + xmlFile +  " /ru " + username + " /f"
    subprocess.call(process, shell=True)
