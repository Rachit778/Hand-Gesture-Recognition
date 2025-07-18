import cv2
import time
import math
import numpy as np
import HandTrackingModule as htm
import pyautogui
import autopy
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# Webcam setup (try 0 if 1 doesn't work)
cap = cv2.VideoCapture(0)
cap.set(3, 640)  # width
cap.set(4, 480)  # height

# Hand detector
detector = htm.handDetector(maxHands=1, detectionCon=0.85, trackCon=0.8)

# Audio setup
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
volRange = volume.GetVolumeRange()
minVol = volRange[0]
maxVol = volRange[1]

# Gesture variables
volBar = 400
volPer = 0
vol = 0
tipIds = [4, 8, 12, 16, 20]
color = (0, 215, 255)
mode = ''
active = 0
hmin, hmax = 50, 200

pyautogui.FAILSAFE = False
pTime = 0

def putText(mode, loc=(250, 450), color=(0, 255, 255)):
    cv2.putText(img, str(mode), loc, cv2.FONT_HERSHEY_COMPLEX_SMALL, 3, color, 3)

while True:
    success, img = cap.read()
    if not success or img is None:
        print("⚠️ Failed to grab frame. Check camera connection or index.")
        continue

    img = detector.findHands(img)
    lmList = detector.findPosition(img, draw=False)
    fingers = []

    if len(lmList) != 0:
        # Thumb
        fingers.append(1 if lmList[tipIds[0]][1] > lmList[tipIds[0] - 1][1] else 0)
        # Fingers
        for id in range(1, 5):
            fingers.append(1 if lmList[tipIds[id]][2] < lmList[tipIds[id] - 2][2] else 0)

        # Detect mode
        if fingers == [0, 0, 0, 0, 0] and active == 0:
            mode = 'N'
        elif fingers in ([0, 1, 0, 0, 0], [0, 1, 1, 0, 0]) and active == 0:
            mode = 'Scroll'
            active = 1
        elif fingers == [1, 1, 0, 0, 0] and active == 0:
            mode = 'Volume'
            active = 1
        elif fingers == [1, 1, 1, 1, 1] and active == 0:
            mode = 'Cursor'
            active = 1

    # SCROLL MODE
    if mode == 'Scroll':
        putText(mode)
        cv2.rectangle(img, (200, 410), (245, 460), (255, 255, 255), cv2.FILLED)
        if len(lmList) != 0:
            if fingers == [0, 1, 0, 0, 0]:
                putText('U', loc=(200, 455), color=(0, 255, 0))
                pyautogui.scroll(300)
            elif fingers == [0, 1, 1, 0, 0]:
                putText('D', loc=(200, 455), color=(0, 0, 255))
                pyautogui.scroll(-300)
            elif fingers == [0, 0, 0, 0, 0]:
                active = 0
                mode = 'N'

    # VOLUME MODE
    if mode == 'Volume':
        putText(mode)
        if len(lmList) != 0:
            if fingers[-1] == 1:
                active = 0
                mode = 'N'
            else:
                x1, y1 = lmList[4][1], lmList[4][2]
                x2, y2 = lmList[8][1], lmList[8][2]
                cx, cy = (x1 + x2) // 2, (y1 + y2) // 2

                # Visual elements
                cv2.circle(img, (x1, y1), 10, color, cv2.FILLED)
                cv2.circle(img, (x2, y2), 10, color, cv2.FILLED)
                cv2.line(img, (x1, y1), (x2, y2), color, 3)
                cv2.circle(img, (cx, cy), 8, color, cv2.FILLED)

                length = math.hypot(x2 - x1, y2 - y1)
                vol = np.interp(length, [hmin, hmax], [minVol, maxVol])
                volBar = np.interp(vol, [minVol, maxVol], [400, 150])
                volPer = np.interp(vol, [minVol, maxVol], [0, 100])
                volume.SetMasterVolumeLevel(vol, None)

                if length < 50:
                    cv2.circle(img, (cx, cy), 11, (0, 0, 255), cv2.FILLED)

                # Volume bar
                cv2.rectangle(img, (30, 150), (55, 400), (209, 206, 0), 3)
                cv2.rectangle(img, (30, int(volBar)), (55, 400), (215, 255, 127), cv2.FILLED)
                cv2.putText(img, f'{int(volPer)}%', (25, 430),
                            cv2.FONT_HERSHEY_COMPLEX, 0.9, (209, 206, 0), 3)

    # CURSOR MODE
    if mode == 'Cursor':
        putText(mode)
        cv2.rectangle(img, (110, 20), (620, 350), (255, 255, 255), 3)
        if fingers[1:] == [0, 0, 0, 0]:
            active = 0
            mode = 'N'
        elif len(lmList) != 0:
            x1, y1 = lmList[8][1], lmList[8][2]
            screenW, screenH = autopy.screen.size()
            X = int(np.interp(x1, [110, 620], [0, screenW]))
            Y = int(np.interp(y1, [20, 350], [0, screenH]))
            X -= X % 2
            Y -= Y % 2
            autopy.mouse.move(X, Y)

            cv2.circle(img, (x1, y1), 7, (255, 255, 255), cv2.FILLED)
            cv2.circle(img, (lmList[4][1], lmList[4][2]), 10, (0, 255, 0), cv2.FILLED)
            if fingers[0] == 0:
                cv2.circle(img, (lmList[4][1], lmList[4][2]), 10, (0, 0, 255), cv2.FILLED)
                pyautogui.click()

    # FPS
    cTime = time.time()
    fps = 1 / ((cTime + 0.01) - pTime)
    pTime = cTime
    cv2.putText(img, f'FPS:{int(fps)}', (480, 50), cv2.FONT_ITALIC, 1, (255, 0, 0), 2)

    cv2.imshow('Hand LiveFeed', img)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break
