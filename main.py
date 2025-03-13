from src import pptxTopng
import os
import cv2
import numpy as np
#
# ppt_file = r"C:\Users\Dhakshesh\Downloads\SmartLiving.pptx"
pptLocation = r"C:\Users\Dhakshesh\Desktop\Daily Contributions\IndexDraw\presentation/"
#
# pptImages = os.listdir(pptLocation)
#
# if pptImages:
#     for i in pptImages:
#         os.remove(os.path.join(pptLocation, i))
#
# pptxTopng.convert_ppt_to_jpg(ppt_file, pptLocation)
# pptImages = os.listdir(pptLocation)
# print("Conversion complete!")


width,height = int(1024*0.7),int(768*0.5)


cap = cv2.VideoCapture(0)
cap.set(3,width)
cap.set(4,height)

pptImages = os.listdir(pptLocation)

#custom lambda method to sort files
sorted_pptImages = sorted(pptImages, key=lambda x: int(x.replace(".png","").split("_")[1]))

#define start slide
startSlide = 5

#setting window size
#want the image to be inside a 700x700 box
desiredSize = (700,700)
#we now create the background for our slides to fit the new window size
canvas = np.zeros((desiredSize[1],desiredSize[0],3),dtype=np.uint8)

while True:
    success,img = cap.read()
    pathFullImage = os.path.join(pptLocation,sorted_pptImages[startSlide])
    currentSlide = cv2.imread(pathFullImage)
    # currentSlide = cv2.resize(currentSlide,(width,height))  #resizing jus the ppt
    #get the og dimensions of the slide
    h,w,_ = currentSlide.shape

    #resize image without affecting aspect ratio by calculating new proportions for width and height
    scale = min(desiredSize[0]/w,desiredSize[1]/h)
    new_w,new_h = int(w*scale), int(h*scale)
    currentSlide = cv2.resize(currentSlide, (new_w, new_h))

    #calculate padding and set the image in the center
    x_offset = (desiredSize[0]-new_w)//2
    y_offset = (desiredSize[1]-new_h)//2
    #
    canvas[y_offset:y_offset+new_h,x_offset:x_offset+new_w] = currentSlide

    cv2.imshow("Image",img)
    #cv2.imshow("Slide",currentSlide)
    #using canvas
    cv2.imshow("Slide", canvas)
    key = cv2.waitKey(1)
    if key == ord('q'):
        break


