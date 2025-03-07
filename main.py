from src import pptxTojpg

ppt_file = r"{source}"
jpg_file = r"{target}"

pptxTojpg.convert_ppt_to_jpg(ppt_file, jpg_file)
print("Conversion complete!")

