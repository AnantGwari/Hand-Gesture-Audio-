import numpy as np
import math
import time
import cv2
import mediapipe as mp
from ctypes import cast,POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities,IAudioEndpointVolume

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_,CLSCTX_ALL,None)
volume = cast(interface,POINTER(IAudioEndpointVolume))
Range = volume.GetVolumeRange()
minVol = Range[0]
maxVol = Range[1]
cap = cv2.VideoCapture(0)
mpHands = mp.solutions.hands
hands = mpHands.Hands(False,2,1,0.5,0.5)
mpDraw = mp.solutions.drawing_utils
pTime =0
cTime =0
fps=10.0
length =0
while True:
    suc,img = cap.read()
    imgRGB = cv2.cvtColor(img,cv2.COLOR_BGR2RGB)
    results = hands.process(imgRGB)
    lmList=[]
    if results.multi_hand_landmarks:
        for handLms in results.multi_hand_landmarks:
            mpDraw.draw_landmarks(img,handLms)
            mpDraw.draw_landmarks(img,handLms,mpHands.HAND_CONNECTIONS)
            for id,lm in enumerate(handLms.landmark):
                h,w,c = img.shape
                cx,cy = int(lm.x*w),int(lm.y*h)
                lmList.append([id,cx,cy])
    if len(lmList)!=0:
        x1,y1 = lmList[4][1],lmList[4][2]
        x2,y2 = lmList[8][1],lmList[8][2]
        cv2.line(img,(x1,y1),(x2,y2),(0,255,0),3)
        length = math.hypot(x2-x1,y2-y1)
    vol = np.interp(length,[50,200],[minVol,maxVol])
    volume.SetMasterVolumeLevel(vol,None)
    cTime = time.time()
    fps = 1/(cTime-pTime)
    pTime = cTime
    cv2.putText(img, str(int(fps)),(10,100), cv2.FONT_HERSHEY_COMPLEX, 1, (0,0,255),2,cv2.LINE_AA)
    cv2.imshow('frame',img)        
    q = cv2.waitKey(1)
    if q ==ord(' ') or q==27:
       break
cap.release()
cv2.destroyAllWindows()
 