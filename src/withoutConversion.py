import win32com.client as win32
import cv2
import numpy as np
import mss
import pygetwindow as gw
import time
def get_ppt_window():
    """Finds the PowerPoint Slide Show window."""
    for window in gw.getWindowsWithTitle("PowerPoint Slide Show"):
        return window
    return None

def capture_and_display_ppt():
    """Captures PowerPoint Slide Show window and displays it using OpenCV."""
    ppt_window = get_ppt_window()

    if not ppt_window:
        print("PowerPoint Slide Show not found! Please start the slide show.")
        return

    # Get window coordinates
    left, top, width, height = ppt_window.left, ppt_window.top, ppt_window.width, ppt_window.height

    # Set up mss to capture the screen in real-time
    with mss.mss() as sct:
        while True:
            # Capture the PowerPoint window screenshot
            screenshot = sct.grab({"left": left, "top": top, "width": width, "height": height})
            frame = np.array(screenshot)
            frame = cv2.cvtColor(frame, cv2.COLOR_BGRA2BGR)  # Convert to BGR for OpenCV

            # Display the captured frame in OpenCV window
            cv2.imshow("PowerPoint Livestream", frame)

            # Wait for a key press to check if it's 'Esc' to exit
            key = cv2.waitKey(1)
            if key == 27:  # Press 'Esc' to exit
                break

    cv2.destroyAllWindows()  # Close OpenCV window after exit

def withoutConversion(pptPath):
    app = win32.Dispatch("PowerPoint.Application")
    objCOM = app.Presentations.Open(FileName=pptPath,WithWindow=1)

    #start the slideshow
    objCOM.SlideShowSettings.Run()
    while True:
        userinput = input("Press d for next Slide")
        if(userinput=='d'):
            objCOM.SlideShowWindow.View.Next()
        elif(userinput=='a'):
            objCOM.SlideShowWindow.View.Previous()

if __name__=='__main__':
    withoutConversion(r"C:\Users\Dhakshesh\Downloads\SmartLiving.pptx")
    time.sleep(20)  # Give time to switch to Slide Show
    capture_and_display_ppt()