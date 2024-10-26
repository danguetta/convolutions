import js
from PIL import Image
import base64
import io
import numpy as np
import openpyxl
import requests
import asyncio

# Save the Excel workbook details
workbook_url = 'https://convolutions.guetta.com/convolutions.xlsm'
first_row = 18
first_col = 11

# Define some global variables we'll use throughout
canvas = js.document.getElementById("canvas")
ctx = canvas.getContext("2d")

# Keep track of the clicks
selection_start = None
selection_end = None
mouse_clicked = False

# Prepare a variable for the image
image = None

# Specify the size of the maximum dimension of the image we want to display
max_dimension = 500

async def upload_image(event):
    '''
    This function is triggered when the "upload image" button is clicked
    '''
    global image

    try:
        file = await (await js.window.showOpenFilePicker())[0].getFile()
        
        if file.name.split('.')[-1] not in ['png', 'jpg', 'jpeg']:
            js.alert("Please choose a PNG or JPG file")
            return

        image = js.Image.new()
        image.src = js.URL.createObjectURL(file)

        # Once the image has loaded, draw it on the canvas
        image.onload = lambda e: display_image(image)
    except:
        pass

async def download_excel(event):
    '''
    This function is triggered when the "download" button is clicked
    '''
    
    # Show the overlay
    # ----------------
    js.document.getElementById("loading_message").innerHTML = "Processing image..."
    js.document.getElementById("loadingOverlay").style.display = "flex"

    await asyncio.sleep(0.1)

    # Crop the image
    # --------------
    x, y, side = draw_selection(draw=False)

    cropped_canvas = js.document.createElement("canvas")
    cropped_canvas.width = side
    cropped_canvas.height = side
    cropped_ctx = cropped_canvas.getContext("2d")
    cropped_ctx.drawImage(canvas, x, y, side, side, 0, 0, side, side)
    
    # Prepare the image
    # -----------------

    final_im = Image.open(io.BytesIO(base64.b64decode(cropped_canvas.toDataURL("image/png").split(',')[1])))

    # Convert the image of black and white
    final_im = final_im.convert('L')

    # Resize to 200 x 194
    final_im = final_im.resize((194, 200))

    # Convert to a numpy array
    final_im = 255 - np.array(final_im)

    print('test1')

    # Load the Excel file, and put the image in it
    # --------------------------------------------
    wb = openpyxl.load_workbook(io.BytesIO(requests.get(workbook_url).content), keep_vba=True)
    
    print('test2')

    print(final_im.shape)
    print(wb['Sheet1']['K18'].value)

    for i in range(200):
        for j in range(194):
            wb['Sheet1'][openpyxl.utils.get_column_letter(first_col + j) + str(first_row + i)] = final_im[i, j]

    print(wb['Sheet1']['K18'].value)

    # Save the Excel file
    # -------------------
    output_file = io.BytesIO()
    wb.save(output_file)
    output_file.seek(0)
    
    binary_data = output_file.read()
    uint8_array = js.Uint8Array.new(binary_data)

    blob = js.Blob.new([uint8_array], { "type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
    url = js.URL.createObjectURL(blob)

    # Create and trigger the download link
    link = js.document.createElement("a")
    link.href = url
    link.download = "convolutions_personalized.xlsm"
    link.click()
    js.URL.revokeObjectURL(url)  # Clean up the URL after download

    # Hide the overlay
    draw_selection()
    js.document.getElementById("loadingOverlay").style.display = "none"

def display_image(img):
    '''
    Displays the image on the canvas, scaled not to be too big, and draws a red
    cropping square on it
    '''

    # Get the height and width of the image and figure out the scaling factor required
    width, height = img.width, img.height

    if width > height:
        scale_factor = max_dimension / width
    else:
        scale_factor = max_dimension / height

    # Calculate the new dimensions
    new_width = int(width * scale_factor)
    new_height = int(height * scale_factor)
    
    # Resize the canvas and display the image
    canvas.width = new_width
    canvas.height = new_height
    ctx.clearRect(0, 0, new_width, new_height)
    ctx.drawImage(img, 0, 0, new_width, new_height)

    # Display the canvas
    js.document.getElementById("image_section").style = "display:inline"

#######################################################
#   Create functions to handle drawing on the image   #
#######################################################

def handle_mouse_down(event):
    global selection_start, selection_end, mouse_clicked

    selection_start = (event.offsetX, event.offsetY)
    selection_end = None
    mouse_clicked = True

def handle_mouse_move(event):
    global selection_start, selection_end, mouse_clicked

    if mouse_clicked:
        selection_end = (event.offsetX, event.offsetY)
        draw_selection()

def handle_mouse_up(event):
    global selection_start, selection_end, mouse_clicked

    if mouse_clicked:
        js.document.getElementById("download_section").style = "display:inline"

    mouse_clicked = False

def draw_selection(draw=True):
    global selection_start, selection_end, mouse_clicked
    
    # Reset canvas and draw image
    display_image(image)

    if selection_start and selection_end:
        x_start, y_start = selection_start
        x_end, y_end = selection_end

        # Figure out the side length of the square
        side = min(abs(x_end - x_start), abs(y_end - y_start))

        # Figure out the top-left corner
        if x_start < x_end:
            x = x_start
        else:
            x = x_start - side
        
        if y_start < y_end:
            y = y_start
        else:
            y = y_start - side

        # Draw the square
        if draw:
            ctx.strokeStyle = "red"
            ctx.lineWidth = 2
            ctx.strokeRect(x, y, side, side)
    
    # Return (x, y, side)
    return (x, y, side)

# Remove the loading overlay
js.document.getElementById("loadingOverlay").style.display = "none"