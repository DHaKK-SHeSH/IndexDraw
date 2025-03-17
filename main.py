from src import pptxTopng
import os
import cv2
import numpy as np

ppt_file = r"C:\Users\Dhakshesh\Downloads\SmartLiving.pptx"
pptLocation = r"C:\Users\Dhakshesh\Desktop\Daily Contributions\IndexDraw\presentation/"

pptImages = os.listdir(pptLocation)

# if pptImages:
#     for i in pptImages:
#         os.remove(os.path.join(pptLocation, i))
#
# pptxTopng.convert_ppt_to_jpg(ppt_file, pptLocation)
# pptImages = os.listdir(pptLocation)
# print("Conversion complete!")


width,height = [2560 ,1440]


cap = cv2.VideoCapture(0)
cap.set(3,width)
cap.set(4,height)

pptImages = os.listdir(pptLocation)

#custom lambda method to sort files
sorted_pptImages = sorted(pptImages, key=lambda x: int(x.replace(".jpg","").split("_")[1]))

#define start slide
startSlide = 23
hgt, wdt = int(1440*0.2),int(2560*0.2)

# trying to create a transparent webcam feed so the slides are unaffected
alpha = 0.5 #provides 30% opacity to webcam
beta = 1-alpha


while True:
    success,img = cap.read()
    pathFullImage = os.path.join(pptLocation,sorted_pptImages[startSlide])
    currentSlide = cv2.imread(pathFullImage)

    cv2.imshow("Image",img)
    # overlay area
    webcamImg = cv2.resize(img,(wdt,hgt))
    h,w,_ = currentSlide.shape
    # region of interest for opacity
    roi = currentSlide[h-hgt:h,w-wdt:w]
    # blend the images
    blended = cv2.addWeighted(webcamImg,alpha,roi,beta,0)
    currentSlide[h-hgt:h,w-wdt:w] = blended

    cv2.namedWindow("Slide", cv2.WINDOW_NORMAL)
    cv2.setWindowProperty("Slide", cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
    cv2.imshow("Slide", currentSlide)
    key = cv2.waitKey(1)
    if key == ord('q'):
        break


