import pyautogui as bot
bot.click(1803, 16)
medida = 188
for i in range(15):
    medida += 31
    bot.sleep(0.2)
    bot.doubleClick(1338, medida)
    bot.sleep(1.5)
    bot.doubleClick(1262,452)
    bot.sleep(1)################
    bot.hotkey('ctrl', 'c')
    bot.click(527, 458)
    bot.sleep(1)################
    bot.press('down')
    bot.hotkey('ctrl','v')
    bot.sleep(1)################
    bot.click(1497, 491)
    bot.hotkey('alt','f4')
    bot.sleep(1)################
    bot.sleep(0.5)
    bot.press('enter')




