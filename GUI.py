import PySimpleGUI as sg         # ![GCM](https://user-images.githubusercontent.com/98583882/151640896-72a5d68a-10c2-4d40-bf44-7d6dde3e44e2.png)
import json
import io
import base64
import os
from PIL import Image
from PIL import ImageWin
from PIL import ImageSequence
from PIL import ImageDraw
from PIL import ImageFont
from datetime import date
import time
import datetime
import moviepy
import moviepy.editor as mp
import ffmpeg
import win32com.client
import win32com.client as win32
from redmail import EmailSender
from redmail import gmail
from redmail import EmailSender
import win32print
import win32ui
from subprocess import call
import images2gif


today = date.today()
TodaysDate = today.strftime("%d %B %Y")

'''
I searched high and low for solutions to the "extract animated GIF frames in Python"
problem, and after much trial and error came up with the following solution based
on several partial examples around the web (mostly Stack Overflow).
There are two pitfalls that aren't often mentioned when dealing with animated GIFs -
firstly that some files feature per-frame local palettes while some have one global
palette for all frames, and secondly that some GIFs replace the entire image with
each new frame ('full' mode in the code below), and some only update a specific
region ('partial').
This code deals with both those cases by examining the palette and redraw
instructions of each frame. In the latter case this requires a preliminary (usually
partial) iteration of the frames before processing, since the redraw mode needs to
be consistently applied across all frames. I found a couple of examples of
partial-mode GIFs containing the occasional full-frame redraw, which would result
in bad renders of those frames if the mode assessment was only done on a
single-frame basis.
Nov 2012
'''


def analyseImage(path):
    im = Image.open(path)
    results = {
        'size': im.size,
        'mode': 'full',
    }
    try:
        while True:
            if im.tile:
                tile = im.tile[0]
                update_region = tile[1]
                update_region_dimensions = update_region[2:]
                if update_region_dimensions != im.size:
                    results['mode'] = 'partial'
                    break
            im.seek(im.tell() + 1)
    except EOFError:
        pass
    return results


def processImage(path):

    mode = analyseImage(path)['mode']

    im = Image.open(path)

    i = 0
    p = im.getpalette()
    last_frame = im.convert('RGBA')

    try:
        while True:
            print("saving %s (%s) frame %d, %s %s" % (path, mode, i, im.size, im.tile))

            if not im.getpalette():
                im.putpalette(p)

            new_frame = Image.new('RGBA', im.size)
            if mode == 'partial':
                 new_frame.paste(last_frame)

            new_frame.paste(im, (0, 0), im.convert('RGBA'))
            new_frame[0].save("F:\\Database\\GreetingCard\\Decoration.gif", save_all=True, append_images=all_frames[1:], duration=100, loop=0)
#            new_frame.save('%s-%d.png' % (''.join(os.path.basename(path).split('.')[:-1]), i), 'PNG')

            i += 1
            last_frame = new_frame
            im.seek(im.tell() + 1)
    except EOFError:
        pass


def paste_image(image_bg, image_element, cx, cy, w, h, rotate=0, h_flip=False):
    image_bg_copy = image_bg.copy()
    image_element_copy = image_element.copy()
    image_element_copy = image_element_copy.resize(size=(w, h))
    if h_flip:
        image_element_copy = image_element_copy.transpose(Image.FLIP_LEFT_RIGHT)
    image_element_copy = image_element_copy.rotate(rotate, expand=True)
    _, _, _, alpha = image_element_copy.split()
    # image_element_copy's width and height will change after rotation
    w = image_element_copy.width
    h = image_element_copy.height
    x0 = cx #- w // 2
    y0 = cy #- h // 2
    x1 = x0 + w
    y1 = y0 + h
    image_bg_copy.paste(image_element_copy, box=(x0, y0, x1, y1), mask=alpha)
    return image_bg_copy


def convert_to_bytes(file_or_bytes, resize=None):
    '''
    Will convert into bytes and optionally resize an image that is a file or a base64 bytes object.
    Turns into  PNG format in the process so that can be displayed by tkinter
    :param file_or_bytes: either a string filename or a bytes base64 image object
    :type file_or_bytes:  (Union[str, bytes])
    :param resize:  optional new size
    :type resize: (Tuple[int, int] or None)
    :return: (bytes) a byte-string object
    :rtype: (bytes)
    '''
    if isinstance(file_or_bytes, str):
        img = PIL.Image.open(file_or_bytes)
    else:
        try:
            img = PIL.Image.open(io.BytesIO(base64.b64decode(file_or_bytes)))
        except Exception as e:
            dataBytesIO = io.BytesIO(file_or_bytes)
            img = PIL.Image.open(dataBytesIO)

    cur_width, cur_height = img.size
    if resize:
        new_width, new_height = resize
        scale = min(new_height/cur_height, new_width/cur_width)
        img = img.resize((int(cur_width*scale), int(cur_height*scale)), PIL.Image.ANTIALIAS)
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    del img
    return bio.getvalue()

def scale_gif(path, scale, new_path=None):
    gif=Image.open(path)
    if not new_path:
        new_path=path
    old_gif_information = {
        'loop': bool(gif.info.get('loop', 1)),
        'duration': gif.info.get('duration', 40),
        'background': gif.info.get('background', 223),
        'extension': gif.info.get('extension', (b'NETSCAPE2.0')),
        'transparency': gif.info.get('transparency', 223)
    }
    new_frames = get_new_frames(gif, scale)
    save_new_gif(new_frames, old_gif_information, new_path)

def get_new_frames(gif, scale):
    new_frames=[]
    actual_frames = gif.n_frames
    for frame in range(actual_frames):
        gif.seek(frame)
        new_frame = Image.new('RGBA', gif.size)
        new_frame.paste(gif)
        new_frame.thumbnail(scale, Image.ANTIALIAS)
        new_frames.append(new_frame)
    return new_frames

def save_new_gif(new_frames, old_gif_information, new_path):
    new_frames[0].save(new_path,
                       save_all=True,
                       append_images=new_frames[1:],
                       duration=old_gif_information['duration'],
                       loop=old_gif_information['loop'],
                       background=old_gif_information['background'],
                       extension=old_gif_information['extension'] ,
                       transparency=old_gif_information['transparency'])


lilac_theme = {'BACKGROUND': '#7A89B7',
                'TEXT': '#FFFFFF',
                'INPUT': '#ADBCDB',
                'TEXT_INPUT': '#000000',
                'SCROLL': '#ADBCDB',
                'BUTTON': ('black', '#697A99'), #'#8B9CBB'
                'PROGRESS': ('#01826B', '#D0D0D0'),
                'BORDER': 1,
                'SLIDER_DEPTH': 1,
                'PROGRESS_DEPTH': 1}


sg.theme_add_new('My New Theme', lilac_theme)
sg.theme('My New Theme')
sg.user_settings_filename(filename='Settings.json', path='F:\\Database\GreetingCard\\')



with open('F:\\Database\\GreetingCard\\Defaults.json', 'r') as file:
    defaults = json.loads(file.read())  # use `json.loads` to do the reverse
file.close()
print(defaults)


for values in defaults:   # {"values": null}
    D_ANN = defaults['-ANN-']
    D_ETHAN = defaults['-ETHAN-']
    D_AVA = defaults['-AVA-']
    D_RYAN = defaults['-RYAN-']
    D_LIAM = defaults['-LIAM-']
    D_MJ = defaults['-MJ-']
    D_BC = defaults['-BC-']
    D_OTHER = defaults['-OTHER-']

    D_DECPNG = defaults['-DECPNG-']
    D_DECJPG = defaults['-DECJPG-']
    D_DECGIF = defaults['-DECGIF-']

    D_CARDPNG = defaults['-CARDPNG-']
    D_CARDJPG = defaults['-CARDJPG-']

    D_PRINT = defaults['-PRINT-']
    D_USEFUTUREDATE = defaults['-USEFUTUREDATE-']
    D_ADDMUSIC = defaults['-ADDMUSIC-']
    D_ADDEMAIL = defaults['-ADDEMAIL-']
    D_TOWHATSAPP = defaults['-TOWHATSAPP-']

    D_RESIZE = defaults['-RESIZE-']
    D_TRANSPARENCY = defaults['-TRANSPARENCY-']
    D_FLIP = defaults['-FLIP-']
    D_SAVECARD = defaults['-SAVECARD-']
    D_MAKEEMAIL = defaults['-MAKE_EMAIL-']

    D_PRINTFILE = defaults['-PRINTFILE-']
    D_DECORATION = defaults['-DECORATION-']
    D_CARDBASE = defaults['-CARDBASE-']
    D_AUDIOFILE = defaults['-AUDIOFILE-']
    D_FLOWERS = defaults['-FLOWERS-']
    D_FUTUREDATE = defaults['-FUTUREDATE-']
    D_MASK = defaults['-MASK-']
    D_ROTATE = defaults['-ROTATE-']
    D_TO = defaults['-TO-']
    D_MESSAGE = defaults['-MESSAGE-']
    D_DESCRIPTION1 = defaults['-DESCRIPTION1-']
    D_DESCRIPTION2 = defaults['-DESCRIPTION2-']
    D_FROM = defaults['-FROM-']
    D_RECIPIENT = defaults['-RECIPIENT-']
    D_RECIPIENT2 = defaults['-RECIPIENT2-']
    D_RECIPIENT3 = defaults['-RECIPIENT3-']
    D_SUBJECT = defaults['-SUBJECT-']
    D_DATEFONTSIZE = defaults['-DATEFONTSIZE-']
    D_DATESHADOWOFFSET = defaults['-DATESHADOWOFFSET-']
    D_DATESHADOWSIZE = defaults['-DATESHADOWSIZE-']
    D_DATEFONTCOLOR = defaults['-DATEFONTCOLOR-']
    D_DATESHADOWCOLOR = defaults['-DATESHADOWCOLOR-']
    D_DATEFONT = defaults['-DATEFONT-']
    D_ADDRESSFONTSIZE = defaults['-ADDRESSFONTSIZE-']
    D_ADDRESSSHADOWOFFSET = defaults['-ADDRESSSHADOWOFFSET-']
    D_ADDRESSSHADOWSIZE = defaults['-ADDRESSSHADOWSIZE-']
    D_ADDRESSFONTCOLOR = defaults['-ADDRESSFONTCOLOR-']
    D_ADDRESSSHADOWCOLOR = defaults['-ADDRESSSHADOWCOLOR-']
    D_ADDRESSFONT = defaults['-ADDRESSFONT-']
    D_MESSAGEFONTSIZE = defaults['-MESSAGEFONTSIZE-']
    D_MESSAGESHADOWOFFSET = defaults['-MESSAGESHADOWOFFSET-']
    D_MESSAGESHADOWSIZE = defaults['-MESSAGESHADOWSIZE-']
    D_MESSAGEFONTCOLOR = defaults['-MESSAGEFONTCOLOR-']
    D_MESSAGESHADOWCOLOR = defaults['-MESSAGESHADOWCOLOR-']
    D_MESSAGEFONT = defaults['-MESSAGEFONT-']
    D_DESCRIPTION_FONT_SIZE = defaults['-DESCRIPTION_FONT_SIZE-']
    D_DESCRIPTION_SHADOW_OFFSET = defaults['-DESCRIPTION_SHADOW_OFFSET-']
    D_DESCRIPTION_SHADOW_SIZE = defaults['-DESCRIPTION_SHADOW_SIZE-']
    D_DESCRIPTION_FONT_COLOR = defaults['-DESCRIPTION_FONT_COLOR-']
    D_DESCRIPTION_SHADOW_COLOR = defaults['-DESCRIPTION_SHADOW_COLOR-']
    D_DESCRIPTION_FONT = defaults['-DESCRIPTION_FONT-']#{"values": null}

    Description_Font_Image = str(D_DESCRIPTION_FONT)
    Description_Font_Image = Description_Font_Image[0:len(Description_Font_Image) - 4]
    Description_Font_Image = "F:\\Database\\Font\\" + Description_Font_Image + ".png"
    D_FLOWERFONTSIZE = defaults['-FLOWERFONTSIZE-']
    D_FLOWERSHADOWOFFSET = defaults['-FLOWERSHADOWOFFSET-']
    D_FLOWERSHADOWSIZE = defaults['-FLOWERSHADOWSIZE-']
    D_FLOWERFONTCOLOR = defaults['-FLOWERFONTCOLOR-']
    D_FLOWERSHADOWCOLOR = defaults['-FLOWERSHADOWCOLOR-']
    D_FLOWERFONT = defaults['-FLOWERFONT-']#{"values": null}

    FlowerFontImage = str(D_FLOWERFONT)
    FlowerFontImage = FlowerFontImage[0:len(FlowerFontImage) - 4]
    FlowerFontImage = "F:\\Database\\Font\\" + FlowerFontImage + ".png"
    D_DateX = defaults['-DateX-']
    D_DateY = defaults['-DateY-']
    D_ToX = defaults['-ToX-']
    D_ToY = defaults['-ToY-']
    D_MessageX = defaults['-MessageX-']
    D_MessageY = defaults['-MessageY-']
    D_Line1X = defaults['-Line1X-']
    D_Line1Y = defaults['-Line1Y-']
    D_Line2X = defaults['-Line2X-']
    D_Line2Y = defaults['-Line2Y-']
    D_FromX = defaults['-FromX-']
    D_FromY = defaults['-FromY-']
    D_FlowersX = defaults['-FlowersX-']
    D_FlowersY = defaults['-FlowersY-']
    D_DecorationX = defaults['-DecorationX-']
    D_DecorationY = defaults['-DecorationY-']
    D_DecorationW = defaults['-DecorationW-']
    D_DecorationH = defaults['-DecorationH-']
    D_VideoX = defaults['-VideoX-']
    D_VideoY = defaults['-VideoY-']
    D_GIFSpeed = defaults['-GIFSpeed-']

if D_CARDPNG:
    Ext = ".png"
elif D_CARDJPG:
    Ext = ".jpg"
if D_DECPNG:
    DecorationType = ".png"
elif D_DECJPG:
    DecorationType = ".jpg"
elif D_DECGIF:
    DecorationType = ".gif"
else:
    pass


choices1 = ('Allicia.ttf', 'Amelia Sophia.ttf', 'Amertha.ttf', 'Amsterdam.ttf', 'angelina.ttf', 'arial.ttf', 'Awesome Season.ttf',
            'Azalleia Ornaments Free.ttf', 'Bailey.ttf', 'Beautiful Roses Script.ttf', 'Beauty Queen.ttf', 
            'Blacksword.ttf', 'Capetown Signature Slant.ttf', 'Charlie.ttf', 'Chocolate.ttf', 'Dark Twenty.ttf', 'Dayland.ttf',
            'happyBirthday.ttf', 'Highlight.ttf', 'Infinite_Stroke.ttf', 'PARCHM.TTF', 'Snakefangs FREE.ttf')

addresses1 = ('1bbmmcc@gmail.com', 'cindymccurley25@gmail.com','mventer16@gmail.com', 'jnath24@yahoo.com',
             'maartenventer@yahoo.com.au', 'annebelventer@yahoo.com.au', 'irenedupb@hotmail.com',
             'garreza@bigpond.net.au', 'Nickkruger@absamail.co.za', 'nic@nid.co.za', 'susankr@absamail.co.za',
             'susankruger@webmail.co.za', 'stephanvh@gmail.com', 'pmcc2261@gmail.com', 'somebody@gmai', 'someone@yaho')

options = [[sg.Frame('To:', [[sg.Radio('Ann', 'people', default=D_ANN, enable_events=True, key='-ANN-'),
                                            sg.Radio('Ethan', 'people', default=D_ETHAN, enable_events=True, key='-ETHAN-'),
                                            sg.Radio('Ava', 'people', default=D_AVA, enable_events=True, key='-AVA-'),
                                            sg.Radio('Ryan', 'people', default=D_RYAN, enable_events=True, key='-RYAN-'),
                                            sg.Radio('Liam', 'people', default=D_LIAM, enable_events=True, key='-LIAM-')],
                                            [sg.Radio('MJ', 'people', default=D_MJ, enable_events=True, key='-MJ-'),
                                            sg.Radio('BC', 'people', default=D_BC, enable_events=True, key='-BC-'),
                                            sg.Radio('Other', 'people', default=D_OTHER, enable_events=True, key='-OTHER-')]], border_width=1)],

[sg.Frame('Decoration Type:', [[sg.Radio('PNG', 'decoration', default=D_DECPNG, enable_events=True, key='-DECPNG-'),
                                            sg.Radio('JPG', 'decoration', default=D_DECJPG, enable_events=True, key='-DECJPG-'),
                                            sg.Radio('GIF', 'decoration', default=D_DECGIF, enable_events=True, key='-DECGIF-')]], border_width=1)],

[sg.Frame('Card Base Type:', [[sg.Radio('PNG', 'card', default=D_CARDPNG, enable_events=True, key='-CARDPNG-'),
                                            sg.Radio('JPG', 'card', default=D_CARDJPG, enable_events=True, key='-CARDJPG-')]], border_width=1)],

         [sg.Frame('Base File:', [[sg.Checkbox('Print', default=False, enable_events=True, key='-PRINT-'),
                                               sg.Checkbox('Future Date', default=False, key='-USEFUTUREDATE-'),
                                               sg.Checkbox('Add Music', default=False, key='-ADDMUSIC-')],
                                               [sg.Checkbox('Add Email', default=False, key='-ADDEMAIL-'),
                                               sg.Checkbox('To WhatsApp', default=False, key='-TOWHATSAPP-')]])],
        [sg.Frame('Decoration', [[sg.Checkbox('Resize', default=True, key='-RESIZE-'),
                                             sg.Checkbox('Transparency', default=True, key='-TRANSPARENCY-'),
                                             sg.Checkbox('Flip', default=False, key='-FLIP-')],
                                             [sg.Checkbox('Save', default=False, enable_events=True, key='-SAVECARD-'),
                                             sg.Checkbox('Make Email', default=False, enable_events=True, key='-MAKE_EMAIL-')]], title_color='black', border_width=1)],
        [sg.Text('Print:'), sg.InputText(default_text=D_PRINTFILE, size=(40, 1), key='-PRINTFILE-')],
        [sg.Text('Card Base:'), sg.InputText(default_text=D_CARDBASE, size=(40, 1), key='-CARDBASE-')],
        [sg.Text('Decoration:'), sg.InputText(default_text=D_DECORATION, size=(40, 1), key='-DECORATION-')],
        [sg.Text('Audio File:'), sg.InputText(default_text=D_AUDIOFILE, size=(25, 1), key='-AUDIOFILE-')],
        [sg.Text('   Flowers:'), sg.InputText(default_text=D_FLOWERS, size=(14, 1), key='-FLOWERS-'), sg.Text('Future Date:'), sg.InputText(default_text=D_FUTUREDATE, size=(14, 1), key='-FUTUREDATE-')],
        [sg.Text('Transparency Mask:'), sg.InputText(default_text=D_MASK, size=(10, 1), key='-MASK-')],
        [sg.Text('Rotate:'), sg.InputText(default_text=D_ROTATE, size=(10, 1), key='-ROTATE-')]
        ]

choices = [[sg.Frame('Card Options:', layout=options)]]

items_chosen = [[sg.Text("", size=(1, 1), key='options')]]

DecorationFolder = 'F:\\Database\\Decorations\\'
CardBaseFolder = 'F:\\Database\\GreetingCard\\'
FontFolder = 'F:\\Database\\Font\\'
Card = D_CARDBASE + Ext
Decoration = D_DECORATION + DecorationType

images_col = [[sg.Frame('Decoration', [[sg.Image(filename=Decoration, subsample=5, key='-DECOR-')]])], 
              [sg.Frame('Card Base', [[sg.Image(filename=Card, subsample=4)]])],
              [sg.Frame('Previous Card', [[sg.Image(filename='F:\\Database\\GreetingCard\\GreetingCard1.png', subsample=4, key='-PREVIOUS-')]])]]

images_col2 = [[sg.Image(filename=('F:\\Database\\GreetingCard\\GreetingCard' + DecorationType), subsample=1, key='-RESULT-')]]

MessageGroup = [[sg.Frame('Message',
           [[sg.Text('                To:'), sg.InputText(default_text=D_TO, size=(30, 1), key='-TO-')],
           [sg.Text('      Message:'), sg.InputText(default_text=D_MESSAGE, size=(30, 1), key='-MESSAGE-')],
           [sg.Text('Description 1:'), sg.InputText(default_text=D_DESCRIPTION1, size=(30, 1), key='-DESCRIPTION1-')],
           [sg.Text('Description 2:'), sg.InputText(default_text=D_DESCRIPTION2, size=(30, 1), key='-DESCRIPTION2-')],
           [sg.Text('           From:'), sg.InputText(default_text=D_FROM, size=(30, 1), key='-FROM-')],
           [sg.Text('')],
           [sg.Text('              To:'), sg.Combo(addresses1, default_value=D_RECIPIENT, size=(30, 8), key='-RECIPIENT-')],
           [sg.Text('             Cc:'), sg.Combo(addresses1, default_value=D_RECIPIENT2, size=(30, 8), key='-RECIPIENT2-')],
           [sg.Text('           Bcc:'), sg.Combo(addresses1, default_value=D_RECIPIENT3, size=(30, 8), key='-RECIPIENT3-')],
           [sg.Text('      Subject:'), sg.InputText(default_text=D_SUBJECT, size=(30, 1), key='-SUBJECT-')]
           ])]]

DateFontImage = str(D_DATEFONT)
DateFontImage = DateFontImage[0:len(DateFontImage)-4]
DateFontImage = "F:\\Database\\Font\\" + DateFontImage + ".png"

DateFontColumn = [[sg.Frame('Date Font', 
           [[sg.Text('           Size:'), sg.Spin(values=(18, 20, 24, 26, 28, 30, 36, 42, 48, 54, 60, 66, 72, 78), initial_value=D_DATEFONTSIZE, key='-DATEFONTSIZE-')], 
           [sg.Text('     Shadow Offset:'), sg.Spin(values=(1, 2, 3), initial_value=D_DATESHADOWOFFSET, key='-DATESHADOWOFFSET-')],
           [sg.Text('Shadow Size'), sg.Spin(values=(0, 1, 2, 3, 4, 5, 6), initial_value=D_DATESHADOWSIZE, key='-DATESHADOWSIZE-')],
           [sg.Text('Font Color:'), sg.InputText(default_text=D_DATEFONTCOLOR, size=(10, 1), key='-DATEFONTCOLOR-')],
           [sg.Text('Shadow Color:'), sg.InputText(default_text=D_DATESHADOWCOLOR, size=(10, 1), key='-DATESHADOWCOLOR-')],
           [sg.Image(filename=DateFontImage, subsample=3)],
           [sg.Text('Font:'), sg.Combo(choices1, default_value=D_DATEFONT, size=(15, 4), key='-DATEFONT-')]])]]  #len(choices1)    , enable_events=True

AddressFontImage = str(D_ADDRESSFONT)
AddressFontImage = AddressFontImage[0:len(AddressFontImage)-4]
AddressFontImage = "F:\\Database\\Font\\" + AddressFontImage + ".png"

AddressFontColumn = [[sg.Frame('Address Font', 
           [[sg.Text('           Size:'), sg.Spin(values=(18, 20, 24, 26, 28, 30, 36, 42, 48, 54, 60, 66, 72, 78), initial_value=D_ADDRESSFONTSIZE, key='-ADDRESSFONTSIZE-')],
           [sg.Text('     Shadow Offset:'), sg.Spin(values=(1, 2, 3), initial_value=D_ADDRESSSHADOWOFFSET, key='-ADDRESSSHADOWOFFSET-')],
           [sg.Text('Shadow Size'), sg.Spin(values=(0, 1, 2, 3, 4, 5, 6), initial_value=D_ADDRESSSHADOWSIZE, key='-ADDRESSSHADOWSIZE-')],
           [sg.Text('FontColor:'), sg.InputText(default_text=D_ADDRESSFONTCOLOR, size=(10, 1), key='-ADDRESSFONTCOLOR-')],
           [sg.Text('Shadow Color:'), sg.InputText(default_text=D_ADDRESSSHADOWCOLOR, size=(10, 1), key='-ADDRESSSHADOWCOLOR-')],
           [sg.Image(filename=AddressFontImage, subsample=3)],
           [sg.Text('Font:'), sg.Combo(choices1, default_value=D_ADDRESSFONT, size=(15, 4), key='-ADDRESSFONT-')]])]]  #len(choices1)

MessageFontImage = str(D_MESSAGEFONT)
MessageFontImage = MessageFontImage[0:len(MessageFontImage) - 4]
MessageFontImage = "F:\\Database\\Font\\" + MessageFontImage + ".png"

MessageFontColumn = [[sg.Frame('Message Font',
           [[sg.Text('           Size:'), sg.Spin(values=(18, 20, 24, 26, 28, 30, 36, 42, 48, 54, 60, 66, 72, 78), initial_value=D_MESSAGEFONTSIZE, key='-MESSAGEFONTSIZE-')],
           [sg.Text('     Shadow Offset:'), sg.Spin(values=(1, 2, 3), initial_value= D_MESSAGESHADOWOFFSET, key='-MESSAGESHADOWOFFSET-')],
           [sg.Text('Shadow Size'), sg.Spin(values=(0, 1, 2, 3, 4, 5, 6), initial_value=D_MESSAGESHADOWSIZE, key='-MESSAGESHADOWSIZE-')],
           [sg.Text('FontColor:'), sg.InputText(default_text=D_MESSAGEFONTCOLOR, size=(10, 1), key='-MESSAGEFONTCOLOR-')],
            [sg.Text('Shadow Color:'), sg.InputText(default_text=D_MESSAGESHADOWCOLOR, size=(10, 1), key='-MESSAGESHADOWCOLOR-')],
           [sg.Image(filename=MessageFontImage, subsample=3)],
           [sg.Text('Font:'), sg.Combo(choices1, default_value=D_MESSAGEFONT, size=(15, 4), key='-MESSAGEFONT-')]])]]  #len(choices1)

Description_Font_Image = str(D_DESCRIPTION_FONT)
Description_Font_Image = Description_Font_Image[0:len(Description_Font_Image) - 4]
Description_Font_Image = "F:\\Database\\Font\\" + Description_Font_Image + ".png"

Description_Font_Column = [[sg.Frame('Description Font',
           [[sg.Text('           Size:'), sg.Spin(values=(18, 20, 24, 26, 28, 30, 36, 42, 48, 54, 60, 66, 72, 78), initial_value=D_DESCRIPTION_FONT_SIZE, key='-DESCRIPTION_FONT_SIZE-')],
           [sg.Text('     Shadow Offset:'), sg.Spin(values=(1, 2, 3), initial_value=D_DESCRIPTION_SHADOW_OFFSET, key='-DESCRIPTION_SHADOW_OFFSET-')],
           [sg.Text('Shadow Size'), sg.Spin(values=(0, 1, 2, 3, 4, 5, 6), initial_value=D_DESCRIPTION_SHADOW_SIZE, key='-DESCRIPTION_SHADOW_SIZE-')],
           [sg.Text('FontColor:'), sg.InputText(default_text=D_DESCRIPTION_FONT_COLOR, size=(10, 1), key='-DESCRIPTION_FONT_COLOR-')],
           [sg.Text('Shadow Color:'), sg.InputText(default_text=D_DESCRIPTION_SHADOW_COLOR, size=(10, 1), key='-DESCRIPTION_SHADOW_COLOR-')],
           [sg.Image(filename=Description_Font_Image, subsample=3)],
           [sg.Text('Font:'), sg.Combo(choices1, default_value=D_DESCRIPTION_FONT, size=(15, 4), key='-DESCRIPTION_FONT-')]])]]  #len(choices1)

FlowerFontImage = str(D_FLOWERFONT)
FlowerFontImage = FlowerFontImage[0:len(FlowerFontImage) - 4]
FlowerFontImage = "F:\\Database\\Font\\" + FlowerFontImage + ".png"

FlowerFontColumn = [[sg.Frame('Flower Font',
           [[sg.Text('           Size:'), sg.Spin(values=(54, 76, 88, 100, 112, 124, 136, 148, 160, 172, 184, 196, 208, 220, 232), initial_value=D_FLOWERFONTSIZE, key='-FLOWERFONTSIZE-')],
           [sg.Text('     Shadow Offset:'), sg.Spin(values=(1, 2, 3), initial_value=D_FLOWERSHADOWOFFSET, key='-FLOWERSHADOWOFFSET-')],
           [sg.Text('Shadow Size'), sg.Spin(values=(0, 1, 2, 3, 4, 5, 6), initial_value=D_FLOWERSHADOWSIZE, key='-FLOWERSHADOWSIZE-')],
           [sg.Text('FontColor:'), sg.InputText(default_text=D_FLOWERFONTCOLOR, size=(10, 1), key='-FLOWERFONTCOLOR-')],
            [sg.Text('Shadow Color:'), sg.InputText(default_text=D_FLOWERSHADOWCOLOR, size=(10, 1), key='-FLOWERSHADOWCOLOR-')],
           [sg.Image(filename=FlowerFontImage, subsample=3)], 
           [sg.Text('Font:'), sg.Combo(choices1, default_value=D_FLOWERFONT, size=(15, 4), key='-FLOWERFONT-')]])]]  #len(choices1)

SETTINGS_PATH = 'F:/Database/GreetingCard'
DateFontImage = "F:\\Database\\Font\\" + DateFontImage #+ ".png"
AddressFontImage = "F:\\Database\\Font\\" + AddressFontImage #+ ".png"
MessageFontImage = "F:\\Database\\Font\\" + MessageFontImage #+ ".png"
Description_Font_Image = "F:\\Database\\Font\\" + Description_Font_Image #+ ".png"
FlowerFontImage = "F:\\Database\\Font\\" + FlowerFontImage #+ ".png"

layout = [[sg.Column(choices, element_justification='l'), sg.VSeperator(),
     sg.Column(images_col), sg.VSeperator(), sg.Column(images_col2)], #, background_color='#9FB8AD' , background_color='#9FB8AD'
         [[sg.Frame('Text and Decor Position Control', [[
         sg.Frame('Date', [[
         sg.Slider(range=(10, 600), orientation='h', size=(8, 20), default_value=D_DateX, resolution=10, key='-DateX-'),
         sg.Slider(range=(500, 1), orientation='v', size=(5, 20), default_value=D_DateY, resolution=10, key='-DateY-')]]),
         sg.Frame('To:', [[
         sg.Slider(range=(10, 200), orientation='h', size=(8, 20), default_value=D_ToX, resolution=10, key='-ToX-'),
         sg.Slider(range=(200, 1), orientation='v', size=(5, 20), default_value=D_ToY, resolution=10, key='-ToY-')]]),
         sg.Frame('Message', [[
         sg.Slider(range=(1, 500), orientation='h', size=(8, 20), default_value=D_MessageX, resolution=10, key='-MessageX-'),
         sg.Slider(range=(500, 1), orientation='v', size=(5, 20), default_value=D_MessageY, resolution=10, key='-MessageY-')]]),
         sg.Frame('Line 1', [[
         sg.Slider(range=(1, 700), orientation='h', size=(8, 20), default_value=D_Line1X, resolution=10, key='-Line1X-'),
         sg.Slider(range=(500, 1), orientation='v', size=(5, 20), default_value=D_Line1Y, resolution=10, key='-Line1Y-')]]),
         sg.Frame('Line 2', [[
         sg.Slider(range=(1, 700), orientation='h', size=(8, 20), default_value=D_Line2X, resolution=10, key='-Line2X-'),
         sg.Slider(range=(500, 1), orientation='v', size=(5, 20), default_value=D_Line2Y, resolution=10, key='-Line2Y-')]]),
         sg.Frame('From:', [[
         sg.Slider(range=(1, 700), orientation='h', size=(8, 20), default_value=D_FromX, resolution=10, key='-FromX-'),
         sg.Slider(range=(500, 1), orientation='v', size=(5, 20), default_value=D_FromY, resolution=10, key='-FromY-')]]),
         sg.Frame('Flowers', [[
         sg.Slider(range=(1, 700), orientation='h', size=(8, 20), default_value=D_FlowersX, resolution=10, key='-FlowersX-'),
         sg.Slider(range=(500, 1), orientation='v', size=(5, 20), default_value=D_FlowersY, resolution=10, key='-FlowersY-')]]),
             sg.Frame('Decoration', [[
             sg.Frame('Position', [[
             sg.Slider(range=(1, 700), orientation='h', size=(8, 20), default_value=D_DecorationX, resolution=10, key='-DecorationX-'),
             sg.Slider(range=(400, 1), orientation='v', size=(5, 20), default_value=D_DecorationY, resolution=10, key='-DecorationY-')]]),
             sg.Frame('Size', [[
             sg.Slider(range=(1, 400), orientation='h', size=(8, 20), default_value=D_DecorationW, resolution=10, key='-DecorationW-'),
             sg.Slider(range=(1, 400), orientation='v', size=(5, 20), default_value=D_DecorationH, resolution=10, key='-DecorationH-')]]),
             sg.Frame('GIF', [[
             sg.Slider(range=(1, 400), orientation='h', size=(8, 20), default_value=D_VideoX, resolution=10, key='-VideoX-'),
             sg.Slider(range=(1, 400), orientation='v', size=(5, 20), default_value=D_VideoY, resolution=10, key='-VideoY-')]]),
             sg.Frame('GIF', [[
             sg.Slider(range=(50, 300), orientation='v', size=(5, 20), default_value=D_GIFSpeed, resolution=10, key='-GIFSpeed-')]])]]),
         ]])],

         [sg.Column(MessageGroup, element_justification='l', vertical_alignment='top'), sg.Column(DateFontColumn, element_justification='l'), sg.Column(AddressFontColumn, element_justification='l'), 
          sg.Column(MessageFontColumn, element_justification='l'), sg.Column(Description_Font_Column, element_justification='l'), sg.Column(FlowerFontColumn, element_justification='l'), 
          sg.Button(button_text='MakeCard', tooltip='Click to submit this window'), sg.Button(button_text='Exit')]
         ]]


window = sg.Window('Greeting Card Maker', layout, default_element_size=(80, 1), grab_anywhere=True)
while True:
    event, values = window.read()
    #default_values = {'values': values}

    if event == sg.WIN_CLOSED or event == 'Exit':
        with open('F:\\Database\\GreetingCard\\Defaults.json', 'w') as file:
            file.write(json.dumps(values))  # use `json.loads` to do the reverse
            file.close()
        break
    if event in (sg.WIN_CLOSED, 'Exit'):
        with open('F:\\Database\\GreetingCard\\Defaults.json', 'w') as file:
            file.write(json.dumps(values))  # use `json.loads` to do the reverse
            file.close()
        break

    if event == '-ANN-':
        window['-RECIPIENT-'].update(value='annebelventer@yahoo.com.au')
        window['-RECIPIENT2-'].update(value='')
        window['-RECIPIENT3-'].update(value='')
        window['-TO-'].update(value='My liefste Ann,')
        window['-MESSAGE-'].update(value='')
        window['-DESCRIPTION1-'].update(value='')
        window['-DESCRIPTION2-'].update(value='Al my liefde,')
        window['-FROM-'].update(value='Maarten')
        window['-SUBJECT-'].update(value='Mooi wense van Maarten')
        window.refresh()
    elif event == '-ETHAN-':
        window['-RECIPIENT-'].update(value='mventer16@gmail.com')
        window['-RECIPIENT2-'].update(value='jnath24@yahoo.com')
        window['-RECIPIENT3-'].update(value='')
        window['-TO-'].update(value='To:  Ethan,')
        window['-MESSAGE-'].update(value='')
        window['-DESCRIPTION1-'].update(value='')
        window['-DESCRIPTION2-'].update(value='')
        window['-FROM-'].update(value='From: Ouma and Oupa')
        window['-SUBJECT-'].update(value='Greeting Card from Ouma and Oupa')
        window.refresh()
    elif event == '-AVA-':
        window['-RECIPIENT-'].update(value='mventer16@gmail.com')
        window['-RECIPIENT2-'].update(value='jnath24@yahoo.com')
        window['-RECIPIENT3-'].update(value='')
        window['-TO-'].update(value='To:  Ava,')
        window['-MESSAGE-'].update(value='')
        window['-DESCRIPTION1-'].update(value='')
        window['-DESCRIPTION2-'].update(value='')
        window['-FROM-'].update(value='From: Ouma and Oupa')
        window['-SUBJECT-'].update(value='Greeting Card from Ouma and Oupa')
    elif event == '-RYAN-':
        window['-RECIPIENT-'].update(value='mventer16@gmail.com')
        window['-RECIPIENT2-'].update(value='jnath24@yahoo.com')
        window['-RECIPIENT3-'].update(value='')
        window['-TO-'].update(value='To:  Ryan,')
        window['-MESSAGE-'].update(value='')
        window['-DESCRIPTION1-'].update(value='')
        window['-DESCRIPTION2-'].update(value='')
        window['-FROM-'].update(value='From: Ouma and Oupa')
        window['-SUBJECT-'].update(value='Greeting Card from Ouma and Oupa')
    elif event == '-LIAM-':
        window['-RECIPIENT-'].update(value='1bbmmcc@gmail.com')
        window['-RECIPIENT2-'].update(value='cindymccurley25@gmail.com')
        window['-RECIPIENT3-'].update(value='')
        window['-TO-'].update(value='To:  Liam,')
        window['-MESSAGE-'].update(value='')
        window['-DESCRIPTION1-'].update(value='')
        window['-DESCRIPTION2-'].update(value='')
        window['-FROM-'].update(value='From: Ouma and Oupa')
        window['-SUBJECT-'].update(value='Greeting Card from Ouma and Oupa')
    elif event == '-MJ-':
        window['-RECIPIENT-'].update(value='mventer16@gmail.com')
        window['-RECIPIENT2-'].update(value='jnath24@yahoo.com')
        window['-RECIPIENT3-'].update(value='')
        window['-TO-'].update(value='To: Marius, Joan, Ethan, Ava and Ryan,')
        window['-MESSAGE-'].update(value='')
        window['-DESCRIPTION1-'].update(value='')
        window['-DESCRIPTION2-'].update(value='')
        window['-FROM-'].update(value='From: Mom and Mom')
        window['-SUBJECT-'].update(value='Greeting Card from Ouma and Oupa')
    elif event == '-BC-':
        window['-RECIPIENT-'].update(value='1bbmmcc@gmail.com')
        window['-RECIPIENT2-'].update(value='cindymccurley25@gmail.com')
        window['-RECIPIENT3-'].update(value='')
        window['-TO-'].update(value='To: Brett, Cindy and Liam,')
        window['-MESSAGE-'].update(value='')
        window['-DESCRIPTION1-'].update(value='')
        window['-DESCRIPTION2-'].update(value='')
        window['-FROM-'].update(value='From: Mom and Dad')
        window['-SUBJECT-'].update(value='Greeting Card from Mom and Dad')
    elif event == '-OTHER-':
        window['-RECIPIENT-'].update(value='')
        window['-RECIPIENT2-'].update(value='')
        window['-RECIPIENT3-'].update(value='')
        window['-TO-'].update(value='To:  ')
        window['-MESSAGE-'].update(value='')
        window['-DESCRIPTION1-'].update(value='')
        window['-DESCRIPTION2-'].update(value='')
        window['-FROM-'].update(value='From: ')
        window['-SUBJECT-'].update(value='Greeting Card from Maarten and Ann')
    else:
        pass
    window.refresh()

    if event == '-SAVECARD-':
        GC = Image.open("F:\\Database\\GreetingCard\\GreetingCard.png", 'r')
        GC.save("F:\\Database\\GreetingCard\\GreetingCard1.png", format="png")
        today = datetime.datetime.now()
        d = today.strftime("%Y-%m-%d %H%M")
        GC.save("F:\\Database\\Postcards\\" + d + ".png", format="png")
        time.sleep(1)
        window['-SAVECARD-'].update(value=False)
    window.refresh()

    if event == '-PRINT-':
        img = Image.open("F:\\Database\\GreetingCard\\GreetingCard.png")
        if img.mode == "RGBA" or "transparency" in img.info:
            r, g, b, a = img.split()
            img = Image.merge("RGB", (r, g, b))
        img.save("F:\\Database\\GreetingCard\\GreetingCardx.png")

        file_in = "F:\\Database\\GreetingCard\\GreetingCardx.png"
        img = Image.open(file_in)

        file_out = "F:\\Database\\GreetingCard\\GreetingCardx.bmp"
        img.save(file_out)


        photoshop = "C:\\Program Files\\Adobe\Adobe Photoshop CS5.1 (64 Bit)\\Photoshop.exe" 
        GCfile = "F:\\Database\\GreetingCard\\GreetingCard.png"
        printer = win32print.GetDefaultPrinter()
        call([photoshop, GCfile])
        time.sleep(1)
        window['-PRINT-'].update(value=False)

    if event == '-MAKE_EMAIL-':
        # email = EmailSender(host="smtp.office365.com", port=587)
        email = EmailSender(
            host='smtp.office365.com',
            port=587,
            user_name='maartenventer@hotmail.com',
            password='watleesek456'
        )
        email.send(
            subject=D_SUBJECT,
            sender="maartenventer@hotmail.com",
            receivers=[D_RECIPIENT, D_RECIPIENT2, D_RECIPIENT3],
            html="""<h5>{{ to }}</h5>
            {{ my_image }}
            <h5>{{ from }}</h5>""",
            body_images={
                'my_image': 'F:\\Database\\GreetingCard\\GreetingCard.png',
            },
            body_params={
                "to": D_TO,
                "from": D_FROM,
            },
        )
        time.sleep(1)
        window['-MAKE_EMAIL-'].update(value=False)
    window.refresh()

    if event == 'MakeCard':
        DecorationFolder = 'F:\\Database\\Decorations\\'
        CardBaseFolder = 'F:\\Database\\GreetingCard\\'
        ResultCardFolder = 'F:\\Database\\GreetingCard\\'
        FontFolder = 'F:\\Database\\Font\\'
        Decoration = DecorationFolder + values['-DECORATION-'] + ".png"
        Card = 'F:\\Database\\GreetingCard\\' + D_CARDBASE + ".png"
        DateFontImage = "F:\\Database\\Font\\" + DateFontImage 
        AddressFontImage = "F:\\Database\\Font\\" + AddressFontImage 
        MessageFontImage = "F:\\Database\\Font\\" + MessageFontImage 
        Description_Font_Image = "F:\\Database\\Font\\" + Description_Font_Image 
        FlowerFontImage = "F:\\Database\\Font\\" + FlowerFontImage 


        D_ANN = values['-ANN-']
        D_ETHAN = values['-ETHAN-']
        D_AVA = values['-AVA-']
        D_RYAN = values['-RYAN-']
        D_LIAM = values['-LIAM-']
        D_MJ = values['-MJ-']
        D_BC = values['-BC-']
        D_OTHER = values['-OTHER-']

        D_DECPNG = values['-DECPNG-']
        D_DECJPG = values['-DECJPG-']
        D_DECGIF = values['-DECGIF-']

        D_CARDPNG = values['-CARDPNG-']
        D_CARDJPG = values['-CARDJPG-']

        D_PRINT = values['-PRINT-']
        D_USEFUTUREDATE = values['-USEFUTUREDATE-']
        D_ADDMUSIC = values['-ADDMUSIC-']
        D_ADDEMAIL = values['-ADDEMAIL-']
        D_TOWHATSAPP = values['-TOWHATSAPP-']

        D_RESIZE = values['-RESIZE-']
        D_TRANSPARENCY = values['-TRANSPARENCY-']
        D_FLIP = values['-FLIP-']
        D_SAVECARD = values['-SAVECARD-']
        D_MAKEEMAIL = values['-MAKE_EMAIL-']

        D_PRINTFILE = values['-PRINTFILE-']
        D_DECORATION = values['-DECORATION-']
        D_CARDBASE = values['-CARDBASE-']
        D_AUDIOFILE = values['-AUDIOFILE-']
        D_FLOWERS = values['-FLOWERS-']
        D_FUTUREDATE = defaults['-FUTUREDATE-']
        D_MASK = values['-MASK-']
        D_ROTATE: int = int(values['-ROTATE-'])
        D_TO = values['-TO-']
        D_MESSAGE = values['-MESSAGE-']
        D_DESCRIPTION1 = values['-DESCRIPTION1-']
        D_DESCRIPTION2 = values['-DESCRIPTION2-']
        D_FROM = values['-FROM-']
        D_RECIPIENT = values['-RECIPIENT-']
        D_RECIPIENT2 = values['-RECIPIENT2-']
        D_RECIPIENT3 = values['-RECIPIENT3-']
        D_SUBJECT = values['-SUBJECT-']
        D_DATEFONTSIZE = values['-DATEFONTSIZE-']
        D_DATESHADOWOFFSET = values['-DATESHADOWOFFSET-']
        D_DATESHADOWSIZE = values['-DATESHADOWSIZE-']
        D_DATEFONTCOLOR = values['-DATEFONTCOLOR-']
        D_DATESHADOWCOLOR = values['-DATESHADOWCOLOR-']
        D_DATEFONT = values['-DATEFONT-']
#        D_DATEFONT = str(D_DATEFONT)
        D_ADDRESSFONTSIZE = values['-ADDRESSFONTSIZE-']
        D_ADDRESSSHADOWOFFSET = values['-ADDRESSSHADOWOFFSET-']
        D_ADDRESSSHADOWSIZE = values['-ADDRESSSHADOWSIZE-']
        D_ADDRESSFONTCOLOR = values['-ADDRESSFONTCOLOR-']
        D_ADDRESSSHADOWCOLOR = values['-ADDRESSSHADOWCOLOR-']
        D_ADDRESSFONT = values['-ADDRESSFONT-']
        D_MESSAGEFONTSIZE = values['-MESSAGEFONTSIZE-']
        D_MESSAGESHADOWOFFSET = values['-MESSAGESHADOWOFFSET-']
        D_MESSAGESHADOWSIZE = values['-MESSAGESHADOWSIZE-']
        D_MESSAGEFONTCOLOR = values['-MESSAGEFONTCOLOR-']
        D_MESSAGESHADOWCOLOR = values['-MESSAGESHADOWCOLOR-']
        D_MESSAGEFONT = values['-MESSAGEFONT-']
        D_DESCRIPTION_FONT_SIZE = values['-DESCRIPTION_FONT_SIZE-']
        D_DESCRIPTION_SHADOW_OFFSET = values['-DESCRIPTION_SHADOW_OFFSET-']
        D_DESCRIPTION_SHADOW_SIZE = values['-DESCRIPTION_SHADOW_SIZE-']
        D_DESCRIPTION_FONT_COLOR = values['-DESCRIPTION_FONT_COLOR-']
        D_DESCRIPTION_SHADOW_COLOR = values['-DESCRIPTION_SHADOW_COLOR-']
        D_DESCRIPTION_FONT = values['-DESCRIPTION_FONT-']  # {"values": null}

        Description_Font_Image = str(D_DESCRIPTION_FONT)
        Description_Font_Image = Description_Font_Image[2:len(Description_Font_Image) - 6]
        Description_Font_Image = "F:\\Database\\Font\\" + Description_Font_Image + ".png"

        D_FLOWERFONTSIZE = values['-FLOWERFONTSIZE-']
        D_FLOWERSHADOWOFFSET = values['-FLOWERSHADOWOFFSET-']
        D_FLOWERSHADOWSIZE = values['-FLOWERSHADOWSIZE-']
        D_FLOWERFONTCOLOR = values['-FLOWERFONTCOLOR-']
        D_FLOWERSHADOWCOLOR = values['-FLOWERSHADOWCOLOR-']
        D_FLOWERFONT = values['-FLOWERFONT-']  # {"values": null}

        FlowerFontImage = str(D_FLOWERFONT)
        FlowerFontImage = FlowerFontImage[2:len(FlowerFontImage) - 6]
        FlowerFontImage = "F:\\Database\\Font\\" + FlowerFontImage + ".png"

        D_DateX = values['-DateX-']
        D_DateY = values['-DateY-']
        D_ToX = values['-ToX-']
        D_ToY = values['-ToY-']
        D_MessageX = values['-MessageX-']
        D_MessageY = values['-MessageY-']
        D_Line1X = values['-Line1X-']
        D_Line1Y = values['-Line1Y-']
        D_Line2X = values['-Line2X-']
        D_Line2Y = values['-Line2Y-']
        D_FromX = values['-FromX-']
        D_FromY = values['-FromY-']
        D_FlowersX = values['-FlowersX-']
        D_FlowersY = values['-FlowersY-']
        D_DecorationX = values['-DecorationX-']
        D_DecorationY = values['-DecorationY-']
        D_DecorationW = values['-DecorationW-']
        D_DecorationH = values['-DecorationH-']
        D_VideoX = values['-VideoX-']
        D_VideoY = values['-VideoY-']
        D_GIFSpeed = values['-GIFSpeed-']

        Decorwidth = D_VideoX
        Decorheight = D_VideoY
        if D_USEFUTUREDATE:
            TodaysDate = D_FUTUREDATE

        if D_CARDPNG:
            Ext = ".png"
        elif D_CARDJPG:
            Ext = ".jpg"
        if D_DECPNG:
            DecorationType = ".png"
        elif D_DECJPG:
            DecorationType = ".jpg"
        elif D_DECGIF:
            DecorationType = ".gif"
        else:
            pass

        DateFont: str = str(D_DATEFONT)
#        DateFont = DateFont[0:len(DateFont) - 4]
        DateFont = "F:\\Database\\Font\\" + DateFont

        AddressFont: str = str(D_ADDRESSFONT)
#        AddressFont = AddressFont[0:len(AddressFont) - 4]
        AddressFont = "F:\\Database\\Font\\" + AddressFont

        MessageFont: str = str(D_MESSAGEFONT)
#        MessageFont = MessageFont[0:len(MessageFont) - 4]
        MessageFont = "F:\\Database\\Font\\" + MessageFont

        Description_Font: str = str(D_DESCRIPTION_FONT)
#        Description_Font = Description_Font[0:len(Description_Font) - 4]
        Description_Font = "F:\\Database\\Font\\" + Description_Font

        FlowerFont: str = str(D_FLOWERFONT)
#        FlowerFont = FlowerFont[0:len(FlowerFont) - 4]
        FlowerFont = "F:\\Database\\Font\\" + FlowerFont

#        D_ADDRESSSHADOWCOLOR: str = str(D_ADDRESSSHADOWCOLOR)
        D_MESSAGESHADOWCOLOR: str = str(D_MESSAGESHADOWCOLOR)
        D_DATESHADOWCOLOR: str = str(D_DATESHADOWCOLOR)
        D_DESCRIPTION_SHADOW_COLOR: str = str(D_DESCRIPTION_SHADOW_COLOR)
        D_FLOWERSHADOWCOLOR: str = str(D_FLOWERSHADOWCOLOR)

        ASFont = ImageFont.truetype(font=AddressFont, size=int(D_ADDRESSFONTSIZE) + int(D_ADDRESSSHADOWSIZE))
        AFont = ImageFont.truetype(font=AddressFont, size=int(D_ADDRESSFONTSIZE))
        MSFont = ImageFont.truetype(font=MessageFont, size=int(D_MESSAGEFONTSIZE) + int(D_MESSAGESHADOWSIZE))
        MFont = ImageFont.truetype(font=MessageFont, size=int(D_MESSAGEFONTSIZE))
        DSFont = ImageFont.truetype(font=Description_Font, size=int(D_DESCRIPTION_FONT_SIZE) + int(D_DESCRIPTION_SHADOW_SIZE))
        DFont = ImageFont.truetype(font=Description_Font, size=int(D_DESCRIPTION_FONT_SIZE))
        FSFont = ImageFont.truetype(font=FlowerFont, size=int(D_FLOWERFONTSIZE) + int(D_FLOWERSHADOWSIZE))
        FFont = ImageFont.truetype(font=FlowerFont, size=int(D_FLOWERFONTSIZE))
        DateFont = ImageFont.truetype(font=DateFont, size=int(D_DATEFONTSIZE))
        Color: int = int(D_MASK)
        Base = Image.open(D_CARDBASE + Ext)
        #Base = Image.open(D_CARDBASE + Ext)
        Decoration = D_DECORATION + DecorationType
        #Decoration = D_DECORATION + DecorationType
        # &&&&&&&&&&&&&&&&&&&&&&&&&&&&&& PNG &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        if Ext == ".png" and DecorationType == ".png":
            draw = ImageDraw.Draw(Base)
            Decoration = Image.open(Decoration, 'r').convert("RGBA")
            width, height = Decoration.size
            Dwidth = width
            Dheight = height
            Aspect = Dheight/Dwidth
            D_DecorationH: int = int(D_DecorationW * Aspect)
            for x in range(0, int(D_ADDRESSSHADOWOFFSET)):
                draw.text((D_ToX + x, D_ToY + x), D_TO, fill=D_ADDRESSSHADOWCOLOR, font=ASFont)
            draw.text((D_ToX, D_ToY), D_TO, fill=D_ADDRESSFONTCOLOR, font=AFont)
            for x in range(0, int(D_ADDRESSSHADOWOFFSET)):
                draw.text((D_FromX + x, D_FromY + x), D_FROM, fill=D_ADDRESSSHADOWCOLOR, font=ASFont)
            draw.text((D_FromX, D_FromY), D_FROM, fill=D_ADDRESSFONTCOLOR, font=AFont)
            for x in range(0, int(D_MESSAGESHADOWOFFSET)):
                draw.text((D_MessageX + x, D_MessageY + x), D_MESSAGE, fill=D_MESSAGESHADOWCOLOR, font=MSFont)
            draw.text((D_MessageX, D_MessageY), D_MESSAGE, fill=D_MESSAGEFONTCOLOR, font=MFont)
            for x in range(0, int(D_DATESHADOWOFFSET)):
                draw.text((D_DateX + x, D_DateY + x), TodaysDate, fill=D_DATESHADOWCOLOR, font=DateFont)
            draw.text((D_DateX, D_DateY), TodaysDate, fill=D_DATEFONTCOLOR, font=DateFont)
            if D_DESCRIPTION1 != "Null":
                for x in range(0, int(D_DESCRIPTION_SHADOW_OFFSET)):
                    draw.text((D_Line1X+x, D_Line1Y+x), D_DESCRIPTION1, fill=D_DESCRIPTION_SHADOW_COLOR, font=DSFont)
            draw.text((D_Line1X, D_Line1Y), D_DESCRIPTION1, fill=D_DESCRIPTION_FONT_COLOR, font=DFont)
            if D_DESCRIPTION2 != "Null":
                for x in range(0, int(D_DESCRIPTION_SHADOW_OFFSET)):
                    draw.text((D_Line2X + x, D_Line2Y + x), D_DESCRIPTION2, fill=D_DESCRIPTION_SHADOW_COLOR, font=DSFont)
            draw.text((D_Line2X, D_Line2Y), D_DESCRIPTION2, fill=D_DESCRIPTION_FONT_COLOR, font=DFont)
            if D_FLOWERS != "Null":
                for x in range(0, int(D_FLOWERSHADOWOFFSET)):
                    draw.text((D_FlowersX + x, D_FlowersY + x), D_FLOWERS, fill=D_FLOWERSHADOWCOLOR, font=FSFont)
            draw.text((D_FlowersX, D_FlowersY), D_FLOWERS, fill=D_FLOWERFONTCOLOR, font=FFont)
            if D_TRANSPARENCY:
                Decoration = Decoration.convert("RGBA")
                datas = Decoration.getdata()
                newData = []
                for item in datas:
                    if item[0] > Color and item[1] > Color and item[2] > Color:
                        newData.append((Color, Color, Color, 0))
                    else:
                        newData.append(item)
                Decoration.putdata(newData)
                Decoration.save("F:\\Database\\GreetingCard\\Decoration.png", format="png")
                Base = paste_image(Base, Decoration, int(D_DecorationX), int(D_DecorationY), int(D_DecorationW), int(D_DecorationH), rotate=int(D_ROTATE), h_flip=int(D_FLIP))
                Base.save("F:\\Database\\GreetingCard\\GreetingCard.png", format="png")
                time.sleep(1)
                Base.close()
                with open('F:\\Database\\GreetingCard\\Defaults.json', 'w') as file:
                    file.write(json.dumps(values))  # use `json.loads` to do the reverse
                    file.close()
                filename1 = 'F:\\Database\\GreetingCard\\GreetingCard.png'
                window['-RESULT-'].update(filename=filename1)
                window.refresh()
                # &&&&&&&&&&&&&&&&&&&&&&&&&&&&&& JPG &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        elif Ext == ".jpg": # and DecorationType == ".png" or DecorationType == ".jpg"
            draw = ImageDraw.Draw(Base)
            Decoration = Image.open(Decoration, 'r').convert('RGBA')
            width, height = Decoration.size
            Dwidth = Decorwidth
            Dheight = Decorheight
            for x in range(0, int(D_ADDRESSSHADOWOFFSET)):
                draw.text((D_ToX + x, D_ToY + x), D_TO, fill=D_ADDRESSSHADOWCOLOR, font=ASFont)
            draw.text((D_ToX, D_ToY), D_TO, fill=D_ADDRESSFONTCOLOR, font=AFont)
            for x in range(0, int(D_ADDRESSSHADOWOFFSET)):
                draw.text((D_FromX + x, D_FromY + x), D_FROM, fill=D_ADDRESSSHADOWCOLOR, font=ASFont)
            draw.text((D_FromX, D_FromY), D_FROM, fill=D_ADDRESSFONTCOLOR, font=AFont)
            for x in range(0, int(D_MESSAGESHADOWOFFSET)):
                draw.text((D_MessageX + x, D_MessageY + x), D_MESSAGE, fill=D_MESSAGESHADOWCOLOR, font=MSFont)
            draw.text((D_MessageX, D_MessageY), D_MESSAGE, fill=D_MESSAGEFONTCOLOR, font=MFont)
            for x in range(0, int(D_DATESHADOWOFFSET)):
                draw.text((D_DateX + x, D_DateY + x), TodaysDate, fill=D_DATESHADOWCOLOR, font=DateFont)
            draw.text((D_DateX, D_DateY), TodaysDate, fill=D_DATEFONTCOLOR, font=DateFont)
            if D_DESCRIPTION1 != "Null":
                for x in range(0, int(D_DESCRIPTION_SHADOW_OFFSET)):
                    draw.text((D_Line1X + x, D_Line1Y + x), D_DESCRIPTION1, fill=D_DESCRIPTION_SHADOW_COLOR, font=DSFont)
            draw.text((D_Line1X, D_Line1Y), D_DESCRIPTION1, fill=D_DESCRIPTION_FONT_COLOR, font=DFont)
            if D_DESCRIPTION2 != "Null":
                for x in range(0, int(D_DESCRIPTION_SHADOW_OFFSET)):
                    draw.text((D_Line2X + x, D_Line2Y + x), D_DESCRIPTION2, fill=D_DESCRIPTION_SHADOW_COLOR, font=DSFont)
            draw.text((D_Line2X, D_Line2Y), D_DESCRIPTION2, fill=D_DESCRIPTION_FONT_COLOR, font=DFont)
            if D_FLOWERS != "Null":
                for x in range(0, int(D_FLOWERSHADOWOFFSET)):
                    draw.text((D_FlowersX + x, D_FlowersY + x), D_FLOWERS, fill=D_FLOWERSHADOWCOLOR, font=FSFont)
            draw.text((D_FlowersX, D_FlowersY), D_FLOWERS, fill=D_FLOWERFONTCOLOR, font=FFont)
            if MakeTransparent == "True":
                Decoration = Decoration.convert("RGBA")
                datas = Decoration.getdata()
                newData = []
                for item in datas:
                    if item[0] > Color and item[1] > Color and item[2] > Color:
                        newData.append((Color, Color, Color, 0))
                    else:
                        newData.append(item)
                Decoration.putdata(newData)
                Decoration.save("F:\\Database\\GreetingCard\\Decoration.png", format="png")
                Base = paste_image(Base, Decoration, DecorationX, DecorationY, Dwidth, Dheight, rotate=Rotation, h_flip=Flip)
                # Base.show()
            Base.save("F:\\Database\\GreetingCard\\GreetingCard.png", format="png")
            # Base.show()
            # &&&&&&&&&&&&&&&&&&&&&&&&&&&&&& GIF &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        elif DecorationType == ".gif":  # and (Ext == ".png" or Ext == ".jpg")
            Dwidth = Decorwidth
            Dheight = Decorheight
            Decoration = Image.open(D_DECORATION + DecorationType)
            print(D_DECORATION, DecorationType)
            print(Decoration)
            Decoration.save("F:\\Database\\GreetingCard\\Decoration.gif", format="gif")
            draw = ImageDraw.Draw(Base)
            for x in range(0, int(D_ADDRESSSHADOWOFFSET)):
                draw.text((D_ToX + x, D_ToY + x), D_TO, fill=D_ADDRESSSHADOWCOLOR, font=ASFont)
            draw.text((D_ToX, D_ToY), D_TO, fill=D_ADDRESSFONTCOLOR, font=AFont)
            for x in range(0, int(D_ADDRESSSHADOWOFFSET)):
                draw.text((D_FromX + x, D_FromY + x), D_FROM, fill=D_ADDRESSSHADOWCOLOR, font=ASFont)
            draw.text((D_FromX, D_FromY), D_FROM, fill=D_ADDRESSFONTCOLOR, font=AFont)
            for x in range(0, int(D_MESSAGESHADOWOFFSET)):
                draw.text((D_MessageX + x, D_MessageY + x), D_MESSAGE, fill=D_MESSAGESHADOWCOLOR, font=MSFont)
            draw.text((D_MessageX, D_MessageY), D_MESSAGE, fill=D_MESSAGEFONTCOLOR, font=MFont)
            for x in range(0, int(D_DATESHADOWOFFSET)):
                draw.text((D_DateX + x, D_DateY + x), TodaysDate, fill=D_DATESHADOWCOLOR, font=DateFont)
            draw.text((D_DateX, D_DateY), TodaysDate, fill=D_DATEFONTCOLOR, font=DateFont)
            if D_DESCRIPTION1 != "Null":
                for x in range(0, int(D_DESCRIPTION_SHADOW_OFFSET)):
                    draw.text((D_Line1X + x, D_Line1Y + x), D_DESCRIPTION1, fill=D_DESCRIPTION_SHADOW_COLOR, font=DSFont)
            draw.text((D_Line1X, D_Line1Y), D_DESCRIPTION1, fill=D_DESCRIPTION_FONT_COLOR, font=DFont)
            if D_DESCRIPTION2 != "Null":
                for x in range(0, int(D_DESCRIPTION_SHADOW_OFFSET)):
                    draw.text((D_Line2X + x, D_Line2Y + x), D_DESCRIPTION2, fill=D_DESCRIPTION_SHADOW_COLOR, font=DSFont)
            draw.text((D_Line2X, D_Line2Y), D_DESCRIPTION2, fill=D_DESCRIPTION_FONT_COLOR, font=DFont)
            if D_FLOWERS != "Null":
                for x in range(0, int(D_FLOWERSHADOWOFFSET)):
                    draw.text((D_FlowersX + x, D_FlowersY + x), D_FLOWERS, fill=D_FLOWERSHADOWCOLOR, font=FSFont)
            draw.text((D_FlowersX, D_FlowersY), D_FLOWERS, fill=D_FLOWERFONTCOLOR, font=FFont)
            Base.save("F:\\Database\\GreetingCard\\GreetingCard.png")
            Base = Image.open("F:\\Database\\GreetingCard\\GreetingCard.png").convert('RGBA')
            #Decoration = Image.open("F:\\Database\\GreetingCard\\Decoration.gif", 'r')
            width, height = Decoration.size
            print(width, height)
            print(D_VideoX, D_VideoY)
            ratio = width / height
            D_VideoX: int = int(D_VideoX)
            D_VideoY: int = int(D_VideoX/ratio)
            print(D_VideoX, D_VideoY)
            #NewWidth = int(D_VideoX)
            #NewHeight: int = int(NewWidth/ratio)
            #frames = images2gif.readGif(Decoration, False)
            #for frame in frames:
            #    frame.thumbnail((NewWidth, NewHeight), Image.ANTIALIAS)
            #images2gif.writeGif('F:\\Database\\GreetingCard\\Decoration.gif', frames)
            Dwidth: int = int(width + D_DecorationX)
            Dheight: int = int(height + D_DecorationY)
            all_frames = []
            for Decoration in ImageSequence.Iterator(Decoration):
                new_frame = Base.copy()
                Decoration = Decoration.convert('RGBA')
                new_frame.paste(Decoration, (int(D_DecorationX), int(D_DecorationY), Dwidth, Dheight), Decoration) #D_VideoX
                a = all_frames.append(new_frame)
            all_frames[0].save("F:\\Database\\GreetingCard\\GreetingCard.gif", save_all=True, append_images=all_frames[1:], duration=D_GIFSpeed, loop=0)
            filename1 = 'F:\\Database\\GreetingCard\\GreetingCard.gif'
            window['-RESULT-'].update(filename=filename1)
            #&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& MTS &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        elif DecorationType == ".mts":  # &&&&&& and (Ext == ".png" or Ext == ".jpg")
            draw = ImageDraw.Draw(Base)
            Dwidth = Decorwidth
            Dheight = Decorheight
            for x in range(0, int(D_ADDRESSSHADOWOFFSET)):
                draw.text((D_ToX + x, D_ToY + x), D_TO, fill=D_ADDRESSSHADOWCOLOR, font=ASFont)
            draw.text((D_ToX, D_ToY), D_TO, fill=D_ADDRESSFONTCOLOR, font=AFont)
            for x in range(0, int(D_ADDRESSSHADOWOFFSET)):
                draw.text((D_FromX + x, D_FromY + x), D_FROM, fill=D_ADDRESSSHADOWCOLOR, font=ASFont)
            draw.text((D_FromX, D_FromY), D_FROM, fill=D_ADDRESSFONTCOLOR, font=AFont)
            for x in range(0, int(D_MESSAGESHADOWOFFSET)):
                draw.text((D_MessageX + x, D_MessageY + x), D_MESSAGE, fill=D_MESSAGESHADOWCOLOR, font=MSFont)
            draw.text((D_MessageX, D_MessageY), D_MESSAGE, fill=D_MESSAGEFONTCOLOR, font=MFont)
            for x in range(0, int(D_DATESHADOWOFFSET)):
                draw.text((D_DateX + x, D_DateY + x), TodaysDate, fill=D_DATESHADOWCOLOR, font=DateFont)
            draw.text((D_DateX, D_DateY), TodaysDate, fill=D_DATEFONTCOLOR, font=DateFont)
            if D_DESCRIPTION1 != "Null":
                for x in range(0, int(D_DESCRIPTION_SHADOW_OFFSET)):
                    draw.text((D_Line1X + x, D_Line1Y + x), D_DESCRIPTION1, fill=D_DESCRIPTION_SHADOW_COLOR, font=DSFont)
            draw.text((D_Line1X, D_Line1Y), D_DESCRIPTION1, fill=D_DESCRIPTION_FONT_COLOR, font=DFont)
            if D_DESCRIPTION2 != "Null":
                for x in range(0, int(D_DESCRIPTION_SHADOW_OFFSET)):
                    draw.text((D_Line2X + x, D_Line2Y + x), D_DESCRIPTION2, fill=D_DESCRIPTION_SHADOW_COLOR, font=DSFont)
            draw.text((D_Line2X, D_Line2Y), D_DESCRIPTION2, fill=D_DESCRIPTION_FONT_COLOR, font=DFont)
            if D_FLOWERS != "Null":
                for x in range(0, int(D_FLOWERSHADOWOFFSET)):
                    draw.text((D_FlowersX + x, D_FlowersY + x), D_FLOWERS, fill=D_FLOWERSHADOWCOLOR, font=FSFont)
            draw.text((D_FlowersX, D_FlowersY), D_FLOWERS, fill=D_FLOWERFONTCOLOR, font=FFont)

            Base.save("F:\\Database\\GreetingCard\\GreetingCard.png")
            # Base.show()
            Base.close()
        else:
            x = 1234567890
window.close()


'''
for values in defaults:   # {"values": null}
    D_ANN = defaults['-ANN-']
    D_ETHAN = defaults['-ETHAN-']
    D_AVA = defaults['-AVA-']
    D_RYAN = defaults['-RYAN-']
    D_LIAM = defaults['-LIAM-']
    D_MJ = defaults['-MJ-']
    D_BC = defaults['-BC-']
    D_OTHER = defaults['-OTHER-']

    D_DECPNG = defaults['-DECPNG-']
    D_DECJPG = defaults['-DECJPG-']
    D_DECGIF = defaults['-DECGIF-']

    D_CARDPNG = defaults['-CARDPNG-']
    D_CARDJPG = defaults['-CARDJPG-']

    D_PRINT = defaults['-PRINT-']
#    D_FUTUREDATE = defaults['-FUTUREDATE-']
    D_ADDMUSIC = defaults['-ADDMUSIC-']
    D_ADDEMAIL = defaults['-ADDEMAIL-']
    D_TOWHATSAPP = defaults['-TOWHATSAPP-']

    D_RESIZE = defaults['-RESIZE-']
    D_TRANSPARENCY = defaults['-TRANSPARENCY-']
    D_FLIP = defaults['-FLIP-']
    D_SAVECARD = defaults['-SAVECARD-']
    D_MAKEEMAIL = defaults['-MAKE_EMAIL-']

    D_PRINTFILE = defaults['-PRINTFILE-']
    D_DECORATION = defaults['-DECORATION-']
    D_CARDBASE = defaults['-CARDBASE-']
    D_AUDIOFILE = defaults['-AUDIOFILE-']
    D_FLOWERS = defaults['-FLOWERS-']
    D_MASK = defaults['-MASK-']
    D_ROTATE = defaults['-ROTATE-']
    D_TO = defaults['-TO-']
    D_MESSAGE = defaults['-MESSAGE-']
    D_DESCRIPTION1 = defaults['-DESCRIPTION1-']
    D_DESCRIPTION2 = defaults['-DESCRIPTION2-']
    D_FROM = defaults['-FROM-']
    D_RECIPIENT = defaults['-RECIPIENT-']
    D_RECIPIENT2 = defaults['-RECIPIENT2-']
    D_RECIPIENT3 = defaults['-RECIPIENT3-']
    D_SUBJECT = defaults['-SUBJECT-']
    D_DATEFONTSIZE = defaults['-DATEFONTSIZE-']
    D_DATESHADOWOFFSET = defaults['-DATESHADOWOFFSET-']
    D_DATESHADOWSIZE = defaults['-DATESHADOWSIZE-']
    D_DATEFONTCOLOR = defaults['-DATEFONTCOLOR-']
    D_DATESHADOWCOLOR = defaults['-DATESHADOWCOLOR-']
    D_DATEFONT = defaults['-DATEFONT-']
    D_ADDRESSFONTSIZE = defaults['-ADDRESSFONTSIZE-']
    D_ADDRESSSHADOWOFFSET = defaults['-ADDRESSSHADOWOFFSET-']
    D_ADDRESSSHADOWSIZE = defaults['-ADDRESSSHADOWSIZE-']
    D_ADDRESSFONTCOLOR = defaults['-ADDRESSFONTCOLOR-']
    D_ADDRESSSHADOWCOLOR = defaults['-ADDRESSSHADOWCOLOR-']
    D_ADDRESSFONT = defaults['-ADDRESSFONT-']
    D_MESSAGEFONTSIZE = defaults['-MESSAGEFONTSIZE-']
    D_MESSAGESHADOWOFFSET = defaults['-MESSAGESHADOWOFFSET-']
    D_MESSAGESHADOWSIZE = defaults['-MESSAGESHADOWSIZE-']
    D_MESSAGEFONTCOLOR = defaults['-MESSAGEFONTCOLOR-']
    D_MESSAGESHADOWCOLOR = defaults['-MESSAGESHADOWCOLOR-']
    D_MESSAGEFONT = defaults['-MESSAGEFONT-']
    D_DESCRIPTION_FONT_SIZE = defaults['-DESCRIPTION_FONT_SIZE-']
    D_DESCRIPTION_SHADOW_OFFSET = defaults['-DESCRIPTION_SHADOW_OFFSET-']
    D_DESCRIPTION_SHADOW_SIZE = defaults['-DESCRIPTION_SHADOW_SIZE-']
    D_DESCRIPTION_FONT_COLOR = defaults['-DESCRIPTION_FONT_COLOR-']
    D_DESCRIPTION_SHADOW_COLOR = defaults['-DESCRIPTION_SHADOW_COLOR-']
    D_DESCRIPTION_FONT = defaults['-DESCRIPTION_FONT-']#{"values": null}

    Description_Font_Image = str(D_DESCRIPTION_FONT)
    Description_Font_Image = Description_Font_Image[2:len(Description_Font_Image) - 6]
    Description_Font_Image = "F:\\Database\\Font\\" + Description_Font_Image + ".png"
    D_FLOWERFONTSIZE = defaults['-FLOWERFONTSIZE-']
    D_FLOWERSHADOWOFFSET = defaults['-FLOWERSHADOWOFFSET-']
    D_FLOWERSHADOWSIZE = defaults['-FLOWERSHADOWSIZE-']
    D_FLOWERFONTCOLOR = defaults['-FLOWERFONTCOLOR-']
    D_FLOWERSHADOWCOLOR = defaults['-FLOWERSHADOWCOLOR-']
    D_FLOWERFONT = defaults['-FLOWERFONT-']#{"values": null}

    FlowerFontImage = str(D_FLOWERFONT)
    FlowerFontImage = FlowerFontImage[2:len(FlowerFontImage) - 6]
    FlowerFontImage = "F:\\Database\\Font\\" + FlowerFontImage + ".png"
    D_DateX = defaults['-DateX-']
    D_DateY = defaults['-DateY-']
    D_ToX = defaults['-ToX-']
    D_ToY = defaults['-ToY-']
    D_MessageX = defaults['-MessageX-']
    D_MessageY = defaults['-MessageY-']
    D_Line1X = defaults['-Line1X-']
    D_Line1Y = defaults['-Line1Y-']
    D_Line2X = defaults['-Line2X-']
    D_Line2Y = defaults['-Line2Y-']
    D_FromX = defaults['-FromX-']
    D_FromY = defaults['-FromY-']
    D_FlowersX = defaults['-FlowersX-']
    D_FlowersY = defaults['-FlowersY-']
    D_DecorationX = defaults['-DecorationX-']
    D_DecorationY = defaults['-DecorationY-']
    D_DecorationW = defaults['-DecorationW-']
    D_DecorationH = defaults['-DecorationH-']
    D_VideoX = defaults['-VideoX-']
    D_VideoY = defaults['-VideoY-']
    D_GIFSpeed = defaults['-GIFSpeed-']


D_ANN = False
D_ETHAN = True
D_AVA = False
D_RYAN = False
D_LIAM = False
D_MJ = False
D_BC = False
D_OTHER = False
D_DECPNG = True
D_DECJPG = False
D_DECGIF = False
D_CARDPNG = True
D_CARDJPG = False
D_PRINT = False
D_FUTUREDATE = False
D_ADDMUSIC = False
D_ADDEMAIL = False
D_TOWHATSAPP = False
D_RESIZE = True
D_TRANSPARENCY = True
D_FLIP = False
D_SAVECARD = False
D_MAKEEMAIL = False
D_PRINTFILE = "F:\\Database\\GreetingCard\\GreetingCard.png"
D_CARDBASE = "Frozen18"
D_DECORATION = "Ana6"
D_AUDIOFILE = "ff"
D_FLOWERS = "abc"
D_MASK = "247"
D_ROTATE = "0"
D_DateX = 460.0
D_DateY = 20.0
D_ToX = 40.0
D_ToY = 30.0
D_MessageX = 100.0
D_MessageY = 140.0
D_Line1X = 50.0
D_Line1Y = 360.0
D_Line2X = 70.0
D_Line2Y = 210.0
D_FromX = 120.0
D_FromY = 480.0
D_FlowersX = 300.0
D_FlowersY = 400.0
D_DecorationX = 600.0
D_DecorationY = 150.0
D_DecorationW = 270.0
D_DecorationH = 380.0
D_VideoX = 330.0
D_VideoY = 320.0
D_GIFSpeed = 240.0
D_TO = "To:  Ethan"
D_MESSAGE = "Hope you..."
D_DESCRIPTION1 = "...get well..."
D_DESCRIPTION2 = "...soon!!!"
D_FROM = "Oupa"
D_RECIPIENT = "mventer16@gmail.com"
D_RECIPIENT2 = "jnath24@yahoo.com"
D_RECIPIENT3 = ""
D_SUBJECT = "Greetings"
D_DATEFONTSIZE = 24
D_DATESHADOWOFFSET = 2
D_DATESHADOWSIZE = 2
D_DATEFONTCOLOR = "#FFFFFF"
D_DATESHADOWCOLOR = "#000000"
D_DATEFONT = "Beauty Queen.ttf"
D_ADDRESSFONTSIZE = 26
D_ADDRESSSHADOWOFFSET = 1
D_ADDRESSSHADOWSIZE = 1
D_ADDRESSFONTCOLOR = "#FFFFFF"
D_ADDRESSSHADOWCOLOR = "#000000"
D_ADDRESSFONT = "angelina.ttf"
D_MESSAGEFONTSIZE = 54
D_MESSAGESHADOWOFFSET = 1
D_MESSAGESHADOWSIZE = 1
D_MESSAGEFONTCOLOR = "#FFFFFF"
D_MESSAGESHADOWCOLOR = "#000000"
D_MESSAGEFONT = "Amertha.ttf"
D_DESCRIPTION_FONT_SIZE = 66
D_DESCRIPTION_SHADOW_OFFSET = 1
D_DESCRIPTION_SHADOW_SIZE = 1
D_DESCRIPTION_FONT_COLOR = "#FFFFFF"
D_DESCRIPTION_SHADOW_COLOR = "#000000"
D_DESCRIPTION_FONT = "Dark Twenty.ttf"
D_FLOWERFONTSIZE = 54
D_FLOWERSHADOWOFFSET = 1
D_FLOWERSHADOWSIZE = 1
D_FLOWERFONTCOLOR = "#FFFFFF"
D_FLOWERSHADOWCOLOR = "#000000"
D_FLOWERFONT = "Azalleia Ornaments Free.ttf"


'''    

'''
            else:
                Base = paste_image(Base, Decoration, D_DecorationX, D_DecorationY, Dwidth, Dheight, rotate=D_ROTATE, h_flip=D_FLIP)
                Base.paste(Decoration, (D_DecorationX, D_DecorationY, Dwidth, Dheight))
                Base.save("F:\\Database\\GreetingCard\\GreetingCard.png", format="png")
                #Base.save("F:\\Database\\GreetingCard\\GreetingCard.png", format="png")
               # time.sleep(1)
                Base.close()
#                with open('F:\\Database\\GreetingCard\\Defaults.json', 'w') as file:
#                    file.write(json.dumps(default_values))  # use `json.loads` to do the reverse
#                    file.close()
                filename1 = 'F:\\Database\\GreetingCard\\GreetingCard.png'
                window['-RESULT-'].update(filename=filename1)
                window.refresh()

        else:
            Base = paste_image(Base, Decoration, D_DecorationX, D_DecorationY, D_DecorationW, D_DecorationH, rotate=D_ROTATE, h_flip=D_FLIP)
            Base.paste(Decoration, (D_DecorationX, D_DecorationY, D_DecorationW, D_DecorationH))
            Base.save("F:\\Database\\GreetingCard\\GreetingCard.png", format="png")
            window['-RESULT-'].update('F:\\Database\\GreetingCard\\GreetingCard.png')
#        with open('F:\\Database\\GreetingCard\\Defaults.json', 'w') as file:
#            file.write(json.dumps(values))  # use `json.loads` to do the reverse
#            file.close()
'''