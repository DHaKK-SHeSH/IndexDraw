import win32com.client

def convert_ppt_to_jpg(srcFile,tgtLocation):
    Application = win32com.client.Dispatch("PowerPoint.Application")
    try:
        Presentation = Application.Presentations.Open(srcFile)
        for i in range(Presentation.Slides.Count):
            Presentation.Slides[i].Export(tgtLocation +"slide_"+str(i)+".jpg", "PNG")
        Presentation.Close()
    finally:
        Application.Quit()
