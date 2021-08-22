# imports
try : 
    import pytesseract
    from PIL import Image, ImageGrab
    import win32api
    import win32con
    from time import sleep
    import openpyxl
except :
    print("Impossible to import libs.")

# list of the items we want to check the price
items_list = "galet brasillant","pepite"

# search bar and first item coords
hdv_search_bar_pos = (370,190)
hdv_first_item_pos = (700,220)

letter_to_key = {
           'a':0x41,
           'b':0x42,
           'c':0x43,
           'd':0x44,
           'e':0x45,
           'f':0x46,
           'g':0x47,
           'h':0x48,
           'i':0x49,
           'j':0x4A,
           'k':0x4B,
           'l':0x4C,
           'm':0x4D,
           'n':0x4E,
           'o':0x4F,
           'p':0x50,
           'q':0x51,
           'r':0x52,
           's':0x53,
           't':0x54,
           'u':0x55,
           'v':0x56,
           'w':0x57,
           'x':0x58,
           'y':0x59,
           'z':0x5A,
           ' ':0x20
}

workbook = openpyxl.load_workbook(filename = 'Prix_items.xlsx')
print(workbook.sheetnames)

for items in items_list :

    sleep(5)

    # click on search bar and clear text
    win32api.SetCursorPos(hdv_search_bar_pos)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,hdv_search_bar_pos[0],hdv_search_bar_pos[1])
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,hdv_search_bar_pos[0],hdv_search_bar_pos[1])
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,hdv_search_bar_pos[0],hdv_search_bar_pos[1])
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,hdv_search_bar_pos[0],hdv_search_bar_pos[1])
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,hdv_search_bar_pos[0],hdv_search_bar_pos[1])
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,hdv_search_bar_pos[0],hdv_search_bar_pos[1])
    win32api.keybd_event(0x2E, 0, 0,0)
    win32api.keybd_event(0x2E, 0, win32con.KEYEVENTF_KEYUP,0)

    sleep(1)

    # write new item to search
    for letters in items :
        win32api.keybd_event(letter_to_key[letters], 0, 0,0)
        win32api.keybd_event(letter_to_key[letters], 0, win32con.KEYEVENTF_KEYUP,0)

    sleep(1)

    # move cursor to hdv first item
    win32api.SetCursorPos(hdv_first_item_pos)

    sleep(1)

    # click on first item
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,hdv_first_item_pos[0],hdv_first_item_pos[1])
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,hdv_first_item_pos[0],hdv_first_item_pos[1])

    sleep(1)

    prices_screen = ImageGrab.grab(bbox = (820, 240, 1030, 380))

    sleep(1)
  
    string = pytesseract.image_to_string(prices_screen)
    
    s = " ".join(string.split())
    s1 = s.replace(" ", "")
    s2 = s1.replace(".", "")
    s3 = s2.replace("lotde100", " 100 : ")
    s4 = s3.replace("lotde10", " 10 : ")
    s5 = s4.replace("lotde1", " 1 : ")

    print(s5)
