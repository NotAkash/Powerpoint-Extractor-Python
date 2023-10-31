"""
In order to avoid having to manually copy paste notes into my e-notebook for courses with slides upto 60 pages. And thankfully because I get access to .pptx
I decided to make a script that will (hopefully) take text from each slide, append it to a .txt file. 
Version 1.2 will hopefully get better at seperating header/body text
"""
import os
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def extractPPT(pptFile, outputDir="Notes.md"):
    # Load the PowerPoint presentation
    
    presentation = Presentation(pptFile)
    text_file = outputDir
    file = open(text_file,"w",encoding="utf-8")
    file.close()


    # Create a subdirectory for images
    imgDirectory = "images"
    os.makedirs(imgDirectory, exist_ok=True)


    topLine = ""
    # Iterate through slides
    for slideNum, slide in enumerate(presentation.slides):
        with open(text_file, "a", encoding="utf-8") as file:
            #Add two newlines after slide
            file.write("\n") 
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if(slide.shapes.index(shape)==0 and len(slide.shapes)>1 and topLine!=str(paragraph.text)):
                            topLine=str(paragraph.text)
                            file.write(f"**{paragraph.text}:**\n")
                        else:
                            #Add indentation based on paragraph level(for sub bullets).
                            indent = "".join("    " for _ in range(paragraph.level))
                            file.write(indent)
                            for run in paragraph.runs:
                                text = run.text
                                file.write(f"+ {text} ")
                            file.write("\n")
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:  # Check if shape is an image (shape_type 3)
                    print(f"image found in slide {slideNum}.")
                    imgFilePath = f"images/img{slideNum}_{shape.shape_id}.png" #Generate image from slide
                    imgByte = shape.image.blob
                    with open(imgFilePath, "wb") as imgFile:
                        imgFile.write(imgByte)

                    file.write(f"![Image]({imgFilePath})\n")
                 
if __name__ == "__main__":
    pptFile = "W4.pptx"  # Replace with your PowerPoint file
    outputDir = "W4.md"  # Replace with your desired output directory

    extractPPT(pptFile, outputDir)
