import sys
import time
from os import listdir
from os.path import isfile, join
import cv2
import numpy as np
import pyautogui
import random
import platform
import subprocess
import keyboard
import os
import openpyxl
pyautogui.FAILSAFE= True

def main():

    time.sleep(1)
##    print('Number of arguments:', len(sys.argv), 'arguments.')
##    print('Argument List:', str(sys.argv))
    
    Export_file_extension=sys.argv[1]
    
    
    if Export_file_extension == "0":
        extension='STEP (*.stp)'
    elif Export_file_extension == "1":
        extension='pdf (*.pdf)'
    elif Export_file_extension == "2":
        extension='DXF (*.dxf)'
        
    
    Name=''
    Path=''
    
    for i in range(len(sys.argv)):
        if sys.argv[i]=="delimiter_za_pajton_skriptu_01010101":
            delimiter_index=i+1
            break

    Correct_Name = sys.argv[delimiter_index:]
    
    
    for i in range(len(Correct_Name)):

        if i!=0:
            Name=Name + ' ' + Correct_Name[i]
        else:
            Name=Name + Correct_Name[i]
    zabrana='\/:*?|<>,!;@'
    
    for j in zabrana:
        Name = Name.replace(j,'_')
        
    for i in range(2,delimiter_index-1):
        if i>2:
            Path=Path + " " + sys.argv[i]
        else:
            Path=sys.argv[i]

    pos1=trazi_sve('Capture1',0.9,'nasao status bar')
    pos2=trazi_podrucje('Capture2',pos1[0]-10,pos1[1]-10,pos1[0]+160,pos1[1]+100,0.9,'nasao tipku share menu')
    pyautogui.moveTo(pos1[0]+pos2[0]+20, pos1[1]+pos2[1]+20,0.5)
    pyautogui.click()
    pos3=trazi_podrucje('Capture3',pos1[0]-200,pos1[1]-10,pos1[0]+160,pos1[1]+300,0.9,'nasao share menu')
    pos4=trazi_podrucje('Capture4',pos1[0]-200,pos1[1]-10,pos1[0]+160,pos1[1]+300,0.9,'nasao tipku export')
    pyautogui.moveTo(pos1[0]-200+pos4[0]+10, pos1[1]+pos4[1],0.25)
    pyautogui.click()
    pos5=trazi_sve('Capture5',0.9,'nasao export menu')
    pyautogui.moveTo(pos5[0]+200, pos5[1]+75,0.25)
    pyautogui.click()
    time.sleep(0.25)
    pyautogui.click(pos5[0]+200, pos5[1]+75)
    time.sleep(0.125)
    pyautogui.hotkey('ctrl', 'A')
    time.sleep(0.125)
    keyboard.write(extension)
    time.sleep(0.125)
    pyautogui.moveTo(pos5[0]+200, pos5[1]+95,0.25)
    pyautogui.click()
    pyautogui.moveTo(pos5[0]+180, pos5[1]+95,0.25)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'A')
    keyboard.write(Name)
    pyautogui.moveTo(pos5[0]+180, pos5[1]+150,0.25)
    pyautogui.click()
    time.sleep(0.25)
    pyautogui.click(pos5[0]+180, pos5[1]+150)
    time.sleep(0.25)
    pyautogui.hotkey('ctrl', 'A')
    time.sleep(0.25)
    keyboard.write(Path)
    time.sleep(0.25)
    pos6=trazi_podrucje('Capture6',pos5[0],pos5[1],pos5[0]+600,pos5[1]+600,0.9,'nasao kontrole na prozoru export')
    pyautogui.moveTo(pos6[0]+pos5[0]+120, pos6[1]+pos5[1]+20,0.25)
    pyautogui.click()
    time.sleep(0.25)

    


def trazi_podrucje(ime_slike,x1,y1,x2,y2, prec, javi):
    cwd = os.path.dirname(sys.argv[0])
    cwd=cwd.replace('\\','/')
    pos=imagesearcharea(cwd+"/%s.png"%ime_slike, x1, y1 , x2, y2, precision=prec)
    if pos[0] != -1:
        print(javi)
        return pos[0],pos[1]
    else:
        while pos[0] == -1:
            time.sleep(0.5)
            print("nije "+javi)
            pos=imagesearcharea(cwd+"/%s.png"%ime_slike, x1, y1 , x2, y2, precision=prec)
            time.sleep(0.1)
        print(javi)
        return pos[0],pos[1]

def trazi_sve(ime_slike, prec, javi):
    cwd = os.path.dirname(sys.argv[0])
    cwd=cwd.replace('\\','/')
    pos=imagesearch(cwd+"/%s.png"%ime_slike, precision=prec)
    if pos[0] != -1:
        print(javi)
        return pos[0],pos[1]
    else:
        while pos[0] == -1:
            time.sleep(0.5)
            print("nije "+javi)
            pos=imagesearch(cwd+"/%s.png"%ime_slike, precision=prec)
            time.sleep(0.1)
        print(javi)
        return pos[0],pos[1]

is_retina = False
if platform.system() == "Darwin":
    is_retina = subprocess.call("system_profiler SPDisplaysDataType | grep 'retina'", shell=True)

'''
grabs a region (topx, topy, bottomx, bottomy)
to the tuple (topx, topy, width, height)
input : a tuple containing the 4 coordinates of the region to capture
output : a PIL image of the area selected.
'''


def region_grabber(region):
    if is_retina: region = [n * 2 for n in region]
    x1 = region[0]
    y1 = region[1]
    width = region[2] - x1
    height = region[3] - y1

    return pyautogui.screenshot(region=(x1, y1, width, height))


'''
Searchs for an image within an area
input :
image : path to the image file (see opencv imread for supported types)
x1 : top left x value
y1 : top left y value
x2 : bottom right x value
y2 : bottom right y value
precision : the higher, the lesser tolerant and fewer false positives are found default is 0.8
im : a PIL image, usefull if you intend to search the same unchanging region for several elements
returns :
the top left corner coordinates of the element if found as an array [x,y] or [-1,-1] if not
'''


def imagesearcharea(image, x1, y1, x2, y2, precision=0.8, im=None):
    if im is None:
        im = region_grabber(region=(x1, y1, x2, y2))
        if is_retina:
            im.thumbnail((round(im.size[0] * 0.5), round(im.size[1] * 0.5)))
##        im.save('testarea2.png')
##        usefull for debugging purposes, this will save the captured region as "testarea.png"

    img_rgb = np.array(im)
    img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
    template = cv2.imread(image, 0)

    res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
    if max_val < precision:
        return [-1, -1]
    return max_loc


'''
click on the center of an image with a bit of random.
eg, if an image is 100*100 with an offset of 5 it may click at 52,50 the first time and then 55,53 etc
Usefull to avoid anti-bot monitoring while staying precise.
this function doesn't search for the image, it's only ment for easy clicking on the images.
input :
image : path to the image file (see opencv imread for supported types)
pos : array containing the position of the top left corner of the image [x,y]
action : button of the mouse to activate : "left" "right" "middle", see pyautogui.click documentation for more info
time : time taken for the mouse to move from where it was to the new position
'''


def click_image(image, pos, action, timestamp, offset=5):
    img = cv2.imread(image)
    height, width, channels = img.shape
    pyautogui.moveTo(pos[0] + r(width / 2, offset), pos[1] + r(height / 2, offset),
                     timestamp)
    pyautogui.click(button=action)


'''
Searchs for an image on the screen
input :
image : path to the image file (see opencv imread for supported types)
precision : the higher, the lesser tolerant and fewer false positives are found default is 0.8
im : a PIL image, usefull if you intend to search the same unchanging region for several elements
returns :
the top left corner coordinates of the element if found as an array [x,y] or [-1,-1] if not
'''


def imagesearch(image, precision=0.8):
    im = pyautogui.screenshot()
    if is_retina:
        im.thumbnail((round(im.size[0] * 0.5), round(im.size[1] * 0.5)))
    # im.save('testarea.png') useful for debugging purposes, this will save the captured region as "testarea.png"
    img_rgb = np.array(im)
    img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
    template = cv2.imread(image, 0)
    template.shape[::-1]

    res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
    if max_val < precision:
        return [-1, -1]
    return max_loc


'''
Searchs for an image on screen continuously until it's found.
input :
image : path to the image file (see opencv imread for supported types)
time : Waiting time after failing to find the image 
precision : the higher, the lesser tolerant and fewer false positives are found default is 0.8
returns :
the top left corner coordinates of the element if found as an array [x,y] 
'''


def imagesearch_loop(image, timesample, precision=0.8):
    pos = imagesearch(image, precision)
    while pos[0] == -1:
        print(image + " not found, waiting")
        time.sleep(timesample)
        pos = imagesearch(image, precision)
    return pos


'''
Searchs for an image on screen continuously until it's found or max number of samples reached.
input :
image : path to the image file (see opencv imread for supported types)
time : Waiting time after failing to find the image
maxSamples: maximum number of samples before function times out.
precision : the higher, the lesser tolerant and fewer false positives are found default is 0.8
returns :
the top left corner coordinates of the element if found as an array [x,y] 
'''


def imagesearch_numLoop(image, timesample, maxSamples, precision=0.8):
    pos = imagesearch(image, precision)
    count = 0
    while pos[0] == -1:
        print(image + " not found, waiting")
        time.sleep(timesample)
        pos = imagesearch(image, precision)
        count = count + 1
        if count > maxSamples:
            break
    return pos


'''
Searchs for an image on a region of the screen continuously until it's found.
input :
image : path to the image file (see opencv imread for supported types)
time : Waiting time after failing to find the image 
x1 : top left x value
y1 : top left y value
x2 : bottom right x value
y2 : bottom right y value
precision : the higher, the lesser tolerant and fewer false positives are found default is 0.8
returns :
the top left corner coordinates of the element as an array [x,y] 
'''


def imagesearch_region_loop(image, timesample, x1, y1, x2, y2, precision=0.8):
    pos = imagesearcharea(image, x1, y1, x2, y2, precision)

    while pos[0] == -1:
        time.sleep(timesample)
        pos = imagesearcharea(image, x1, y1, x2, y2, precision)
    return pos


'''
Searches for an image on the screen and counts the number of occurrences.
input :
image : path to the target image file (see opencv imread for supported types)
precision : the higher, the lesser tolerant and fewer false positives are found default is 0.9
returns :
the number of times a given image appears on the screen.
optionally an output image with all the occurances boxed with a red outline.
'''


def imagesearch_count(image, precision=0.9):
    img_rgb = pyautogui.screenshot()
    if is_retina:
        img_rgb.thumbnail((round(img_rgb.size[0] * 0.5), round(img_rgb.size[1] * 0.5)))
    img_rgb = np.array(img_rgb)
    img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
    template = cv2.imread(image, 0)
    w, h = template.shape[::-1]
    res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)
    loc = np.where(res >= precision)
    count = 0
    for pt in zip(*loc[::-1]):  # Swap columns and rows
        # cv2.rectangle(img_rgb, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 2) // Uncomment to draw boxes around found occurrences
        count = count + 1
    # cv2.imwrite('result.png', img_rgb) // Uncomment to write output image with boxes drawn around occurrences
    return count


def r(num, rand):
    return num + rand * random.random()




if __name__ == "__main__":
    main()
