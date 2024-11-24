import os
import time
import json

try:
    import win32com.client
except ImportError as e:
    raise ImportError("The 'pywin32' module is not installed. Please install it using 'pip install pywin32'.") from e

# Define files
input_pptx = "template.pptx"
input_sourceData = "sourceData.json"
output_mp4 = "video.mp4"

if not os.path.exists(input_sourceData):
    raise FileNotFoundError(f"The SourceData file at {input_sourceData} does not exist.")

# Read the JSON file into the replacements dictionary
 
with open(input_sourceData, 'r') as file:
    print("Reading Source Data...")
    #print(file.read())
    replacements = json.load(file)

# Ensure paths are valid
if not os.path.exists(input_pptx):
    raise FileNotFoundError(f"The PowerPoint file at {input_pptx} does not exist.")

#Use COM to get the app
ppt_app = win32com.client.Dispatch("PowerPoint.Application")
try:
    input_pptx_path = os.path.join(os.path.dirname(__file__), input_pptx)
    presentation = ppt_app.Presentations.Open(input_pptx_path, WithWindow=False)
    
    time.sleep(2)  # Short delay to ensure it is ready
    
    # Replace text keys with values in each text box on each slide
    for slide in presentation.Slides:
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                for placeholder, replacement in replacements.items():
                    placeholder = f"[{placeholder.split('_')[0]}]"
                    if placeholder in shape.TextFrame.TextRange.Text:
                        shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text.replace(placeholder, str(replacement))

    print("Creating Video...")
    #See https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.createvideo for more info on the CreateVideo method
    output_mp4_path = os.path.join(os.path.dirname(__file__), output_mp4)
    presentation.CreateVideo(output_mp4_path,True,5,720,30,85)
    
    #Had to wait again :(
    time.sleep(1)

    # Loop waiting for the video export to finish
    #0 None
    #1 In Progress
    #2 Queued
    #3 Done
    #4 Failed

    while presentation.CreateVideoStatus == 1:
        print("Still Creating Video...")
        time.sleep(1)


    if presentation.CreateVideoStatus == 3:
        print("Video export completed successfully.")
    elif presentation.CreateVideoStatus == 4:
        print("Video export failed.")
    else:
        print(f"Video export reported a non successful code. Exit code {presentation.CreateVideoStatus}")
    
    presentation.Close()
except Exception as e:
    print(f"An error occurred: {e}")
finally:
    ppt_app.Quit()