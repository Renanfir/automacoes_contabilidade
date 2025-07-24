import pyautogui as bot

#Deixar o primeiro arquivo com o primeiro clique dado
bot.click(1800, 12)
bot.press('f2')
bot.sleep(1)
for i in range(32):
    bot.press('right')
    bot.press('space')
    bot.write("-")
    bot.press('space')
    bot.write('ALVARA')
    bot.press('tab')
    