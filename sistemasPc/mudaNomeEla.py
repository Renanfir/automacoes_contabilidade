import pyautogui as bot
import string
linha = 16
altura = 75
contador = 144
bot.click(1805, 17)
bot.sleep(1.5)
for i in range(18):  
    bot.click(472, altura)
    bot.sleep(0.2)
    bot.click(803, 431)
    bot.hotkey('ctrl','a')
    bot.press('delete')
    bot.write('a'+str(contador))
    bot.sleep(0.3)
    bot.click(1774, 997)
    print(contador)
    contador += 1
    altura += linha

#ELA 105
#ADRIANO 167