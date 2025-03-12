from src import pptxTopng
import os
import cv2

ppt_file = r"C:\Users\Dhakshesh\Downloads\SmartLiving.pptx"
pptLocation = r"C:\Users\Dhakshesh\Desktop\Daily Contributions\IndexDraw\presentation/"

pptImages = os.listdir(pptLocation)

if pptImages:
    for i in pptImages:
        os.remove(os.path.join(pptLocation, i))

pptxTopng.convert_ppt_to_jpg(ppt_file, pptLocation)
pptImages = os.listdir(pptLocation)
print("Conversion complete!")


width,height = [1280, 720]


cap = cv2.VideoCapture(0)
cap.set(3,width)
cap.set(4,height)

pptImages = os.listdir(pptLocation)

#custom lambda method to sort files
sorted_pptImages = sorted(pptImages, key=lambda x: int(x.replace(".png","").split("_")[1]))

#define start slide
startSlide = 0

while True:
    success,img = cap.read()
    pathFullImage = os.path.join(pptLocation,sorted_pptImages[startSlide])
    currentSlide = cv2.imread(pathFullImage)
    currentSlide = cv2.resize(currentSlide, (width, height))
    cv2.imshow("Image",img)
    cv2.imshow("Slide",currentSlide)
    key = cv2.waitKey(1)
    if key == ord('q'):
        break


