import cv2
import mediapipe as mp
import math
import numpy as np

# pycaw library
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

mpDrawing = mp.solutions.drawing_utils
mp_hands = mp.solutions.hands
HandIm = mp_hands.Hands()

# Function to capture video; 0 denotes default camera
cam = cv2.VideoCapture(0)     

while True:
    success, img = cam.read() 
    imgRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)      # Convert img from BGR to RGB 
    results = HandIm.process(imgRGB)
    #print(results.multi_hand_landmarks)

    cv2.rectangle(img, (600,300), (600,450), (255,64,6), 30)

    startV = (600,450); endV = (600, 450 - round(volume.GetMasterVolumeLevelScalar()*150))

    if volume.GetMute() == 0:
        cv2.line(img, startV, endV, (0,255,0), 20)
    else:
        cv2.line(img, startV, endV, (230, 5, 180), 5)
        #cv2.circle(img, (600,450), 3, (230, 5, 180), 5)
        cv2.rectangle(img, (25,200), (600,70), (255,255,255), cv2.FILLED)
        cv2.putText(img, "Your system is muted.", (100, 100), cv2.FONT_HERSHEY_TRIPLEX, 0.95, (255,12,255), 1)
        cv2.putText(img, " Pls unmute your device to be able to hear audio.", (50, 170), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (95,46,255), 1)
        """
        This code shows the volume bar on the screen:
        in green if the device is not muted (Even if the volume is 0)
        in violet if the system has been physically muted by the user.
        (Specifically required for windows 11)
        """


    device_vol = str(round(volume.GetMasterVolumeLevelScalar()*100))
    cv2.putText(img, device_vol, (570, 270), cv2.FONT_HERSHEY_DUPLEX, 1, (170,255,255), 2)


    if results.multi_hand_landmarks:
        for hand_lms in results.multi_hand_landmarks:
            lmList = []
            for id, lm in enumerate(hand_lms.landmark):
                #print(id,lm)      # lm will give x,y,z coordinates of each landmark; id tells us which landmark each 'lm' belongs to 
                h, w, c = img.shape
                #print(h,w,c)
                cx, cy = int(lm.x*w), int(lm.y*h)
                #print(id, cx, cy)
                lmList.append([id, cx, cy])
            #print(lmList)

            if lmList:
                x1, y1 = lmList[4][1], lmList[4][2]
                x2, y2 = lmList[8][1], lmList[8][2]

                cv2.circle(img, (x1,y1), 10, (153, 190, 13), cv2.FILLED)
                cv2.circle(img, (x2,y2), 10, (153, 190, 13), cv2.FILLED)
                cv2.line(img, (x1,y1), (x2,y2), (2,243,0), 4)

                length = math.hypot((x2-x1), (y2-y1))   # Distance formula to calculate length
                #print(length)

                volRange = volume.GetVolumeRange()    
                minVol = volRange[0]
                maxVol = volRange[1]

                vol = np.interp(length, [30,220], [minVol, maxVol])
                volume.SetMasterVolumeLevel(vol, None)     # To set volume as 'vol'         
                
    cv2.imshow("Camera",img)
    cv2.waitKey(1)