import openpyxl,pyperclip,pyautogui

workbook=openpyxl.load_workbook('products.xlsx')
sheet=workbook['Produtos']

for row in sheet.iter_rows(min_row=2):
    #page 1
    product=row[0].value
    pyperclip.copy(product)
    pyautogui.click(590,210,duration=2)
    pyautogui.hotkey('ctrl','v')

    description=row[1].value
    pyperclip.copy(description)
    pyautogui.click(590,330,duration=.5)
    pyautogui.hotkey('ctrl','v')

    category=row[2].value
    pyperclip.copy(category)
    pyautogui.click(590,440,duration=.5)
    pyautogui.hotkey('ctrl','v')

    product_id=row[3].value
    pyperclip.copy(product_id)
    pyautogui.click(590,520,duration=.5)
    pyautogui.hotkey('ctrl','v')
    
    weight=row[4].value
    pyperclip.copy(weight)
    pyautogui.click(590,610,duration=.5)
    pyautogui.hotkey('ctrl','v')
    
    dimensions=row[5].value
    pyperclip.copy(dimensions)
    pyautogui.click(590,695,duration=.5)
    pyautogui.hotkey('ctrl','v')

    #continue-btn
    pyautogui.click(115,745,duration=1)

    #page 2
    price=row[6].value
    pyperclip.copy(price)
    pyautogui.click(590,240,duration=2)
    pyautogui.hotkey('ctrl','v')

    stock=row[7].value
    pyperclip.copy(stock)
    pyautogui.click(590,330,duration=.5)
    pyautogui.hotkey('ctrl','v')

    expiration_date=row[8].value
    pyperclip.copy(expiration_date)
    pyautogui.click(590,415,duration=.5)
    pyautogui.hotkey('ctrl','v')

    color=row[9].value
    pyperclip.copy(color)
    pyautogui.click(590,500,duration=.5)
    pyautogui.hotkey('ctrl','v')

    size=row[10].value
    pyautogui.click(590,590,duration=.5)
    if(size=='Pequeno'):
        pyautogui.click(590,625,duration=.5)
    elif(size=='Grande'):
        pyautogui.click(590,690,duration=.5)
    else:
        pyautogui.click(590,655,duration=.5)

    material=row[11].value
    pyperclip.copy(material)
    pyautogui.click(590,680,duration=.5)
    pyautogui.hotkey('ctrl','v')

    #continue-btn
    pyautogui.click(115,745,duration=.5)

    brand=row[12].value
    pyperclip.copy(brand)
    pyautogui.click(590,260,duration=.5)
    pyautogui.hotkey('ctrl','v')

    origin=row[13].value
    pyperclip.copy(origin)
    pyautogui.click(590,350,duration=.5)
    pyautogui.hotkey('ctrl','v')

    misc=row[14].value
    pyperclip.copy(misc)
    pyautogui.click(590,455,duration=.5)
    pyautogui.hotkey('ctrl','v')

    barcode=row[15].value
    pyperclip.copy(barcode)
    pyautogui.click(590,575,duration=.5)
    pyautogui.hotkey('ctrl','v')

    location=row[16].value
    pyperclip.copy(location)
    pyautogui.click(590,650,duration=.5)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(115,715,duration=1)

    pyautogui.click(560,490,duration=1)

    pyautogui.click(395,500,duration=1)