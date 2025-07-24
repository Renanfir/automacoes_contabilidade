import pyautogui as bot
import pyperclip
def print_ascii_art():
    print("""
███████ ██    ██ ███    ██  ██████ ██  ██████  ███    ██  ██████  ██     ██
██      ██    ██ ████   ██ ██      ██ ██    ██ ████   ██ ██    ██ ██     ██
█████   ██    ██ ██ ██  ██ ██      ██ ██    ██ ██ ██  ██ ██    ██ ██     ██
██      ██    ██ ██  ██ ██ ██      ██ ██    ██ ██  ██ ██ ██    ██ ██     ██
██       ██████  ██   ████  ██████ ██  ██████  ██   ████  ██████   ███████

    """)

print_ascii_art()   

bot.click(1805,16)
bot.sleep(0.3)
bot.click(396,796)
bot.hotkey("ctrl","a")
bot.hotkey("ctrl","c")
bot.press('tab')
email = pyperclip.paste()
if(email.islower()):
    ...
else:
    email = str(email).lower()
    
bot.write(email)
bot.press('tab')
bot.press('tab')
bot.press('enter')
bot.sleep(0.3)
bot.press('tab')
bot.press("right")
bot.press('tab')
bot.write('') 
bot.press('tab')
bot.write("")
bot.press("tab")
bot.write("")

bot.press("tab")
bot.press('space')
bot.press("tab")
bot.press("tab")
bot.press("tab")
bot.press('enter')
bot.sleep(0.3)
bot.press("tab")
bot.press("tab")
bot.press('space')
bot.press("tab")
bot.press("tab")
bot.press('space')
bot.press("tab")
bot.press('space')
bot.press("tab")
bot.press('space')
bot.press("tab")
bot.press('space')
bot.press("tab")
bot.press("tab")
bot.press("tab")
bot.press("enter")
bot.press("tab")
bot.write("27112029")
bot.press("tab")
bot.press("tab")
bot.press("tab")
bot.press("enter")
bot.sleep(0.4)
for i in range(14):
    bot.press("tab")
bot.press('enter')










