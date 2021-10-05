import cv2
import numpy as npcmd
import face_recognition
import os
from datetime import datetime

# from PIL import ImageGrab


def markAttendance(name):
    with open('Attendance.csv', 'r+') as f:
        myDataList = f.readlines()
        return myDataList

x = markAttendance("name")
print(x)