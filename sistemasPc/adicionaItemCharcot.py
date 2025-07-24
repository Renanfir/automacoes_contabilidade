import pyautogui as bot
qtdLinhas = 15 #PADRÃO = 25
primeira = 95 #PADRÃO = 95



bot.click(1805, 18)
bot.sleep(1.5)
for i in range(qtdLinhas):
    bot.sleep(0.5)
    bot.click(529, primeira)
    bot.sleep(0.2)
    bot.click(241, 55)
    bot.sleep(1.2)
    bot.write('23')
    bot.press('enter')
    bot.press('enter')
    bot.press('enter')
    bot.press('enter')
    bot.write('3334')
    bot.press('enter')
    bot.press('enter')
    bot.write('1')
    bot.click(1189, 897)
    bot.sleep(0.2)
    bot.click(1273, 897)
    primeira += 15



#TABELA DE MEDIDAS
#               1 - 95
#               2 - 110
#               3 - 125
#               4 - 140
#               5 - 155
#               6 - 170
#               7 - 185
#               8 - 200
#               9 - 215
#               10 - 230
#               11 - 245
#               12 - 260
#               13 - 275
#               14 - 290
#               15 - 305
#               16 - 320
#               17 - 335
#               18 - 350
#               19 - 365
#               20 - 380
#               21 - 395
#               22 - 410
#               23 - 425
#               24 - 440
#               25 - 455
#               26 - 470
#               27 - 485