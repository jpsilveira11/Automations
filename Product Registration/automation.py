# Destination: https://cadastro-produtos-devaprender.netlify.app
#dependencies: openpyxl
import openpyxl,pyperclip,pyautogui
# Steps
# 1. Open xlsx
workbook=openpyxl.load_workbook('products.xlsx')
sheet=workbook['Produtos']
# 2. get a field's data
for row in sheet.iter_rows(min_row=2):
    #page 1
    product=row[0].value
    pyperclip.copy(product)
    pyautogui.click(349,220,duration=2)
    pyautogui.hotkey('ctrl','v')

    description=row[1].value
    pyperclip.copy(description)
    pyautogui.click(388,323,duration=.5)
    pyautogui.hotkey('ctrl','v')

    category=row[2].value
    pyperclip.copy(category)
    pyautogui.click(397,429,duration=.5)
    pyautogui.hotkey('ctrl','v')

    product_id=row[3].value
    pyperclip.copy(product_id)
    pyautogui.click(403,526,duration=.5)
    pyautogui.hotkey('ctrl','v')
    
    weight=row[4].value
    pyperclip.copy(weight)
    pyautogui.click(407,605,duration=.5)
    pyautogui.hotkey('ctrl','v')
    
    dimensions=row[5].value
    pyperclip.copy(dimensions)
    pyautogui.click(412,691,duration=.5)
    pyautogui.hotkey('ctrl','v')

    #continue-btn
    pyautogui.click(152,745,duration=1)

    #page 2
    price=row[6].value
    pyperclip.copy(price)
    pyautogui.click(381,244,duration=2)
    pyautogui.hotkey('ctrl','v')

    stock=row[7].value
    pyperclip.copy(stock)
    pyautogui.click(422,321,duration=.5)
    pyautogui.hotkey('ctrl','v')

    expiration_date=row[8].value
    pyperclip.copy(expiration_date)
    pyautogui.click(411,412,duration=.5)
    pyautogui.hotkey('ctrl','v')

    color=row[9].value
    pyperclip.copy(color)
    pyautogui.click(430,498,duration=.5)
    pyautogui.hotkey('ctrl','v')

    size=row[10].value
    pyautogui.click(411,581,duration=.5)
    if(size=='Pequeno'):
        pyautogui.click(462,618,duration=.5)
    elif(size=='Grande'):
        pyautogui.click(357,660,duration=.5)
    else:
        pyautogui.click(347,639,duration=.5)

    material=row[11].value
    pyperclip.copy(material)
    pyautogui.click(438,680,duration=.5)
    pyautogui.hotkey('ctrl','v')

    #continue-btn
    pyautogui.click(152,745,duration=.5)

    brand=row[12].value
    origin=row[13].value
    misc=row[14].value
    barcode=row[15].value
    location=row[16].value
# 3. paste it on the destination
# 4. repeat until the row is finished
# 5. repeat until sheet is finished
# 6. confirm the 1st pop-up
# 7. start the process over