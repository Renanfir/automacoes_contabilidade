import pyautogui as bot
padrao = 120

for i in range(10):
    bot.click(572, padrao)
    bot.sleep(0.5)
    bot.click(529, 569)
    bot.doubleClick(498, 532)
    bot.click(1185, 331)
    bot.write("NFE")
    bot.click(823, 994)
    padrao += 20

