"""
In order to avoid having to manually copy paste notes into my e-notebook for courses with slides upto 60 pages. And thankfully because I get access to .pptx
I decided to make a script that will (hopefully) take text from each slide, append it to a .txt file. 
Version 1.2 will hopefully get better at seperating header/body text
"""

from pptx import Presentation

def extractPPT(pptFile, outputDir):
    # Load the PowerPoint presentation
    
    presentation = Presentation(pptFile)
    text_file = f"{outputDir}"
    
    # Iterate through slides
    for slideNum, slide in enumerate(presentation.slides):
        with open(text_file, "a", encoding="utf-8") as file:
            #Add two newlines after slide
            file.write("\n\n") 
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            text = run.text
                            file.write(f"{text} ")
                            
                        # Add a newline after each paragraph
                        file.write("\n")

if __name__ == "__main__":
    pptFile = "W4.pptx"  # Replace with your PowerPoint file
    outputDir = "W4.txt"  # Replace with your desired output directory

    extractPPT(pptFile, outputDir)
