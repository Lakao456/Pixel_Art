import os
import pip

try:
    import xlsxwriter
    from PIL import Image
except ImportError:
    pip.main(['install', 'pillow'])
    pip.main(['install', 'xlsxwriter'])
    import xlsxwriter
    from PIL import Image

for image in os.listdir(os.getcwd()+'\\Input'):
    print(image)
    im = Image.open("Input/%s" % image)
    colours = []

    workbook = xlsxwriter.Workbook('Output_coloured\\%s_PixelArt_COLOURED.xlsx' % image[:-4])
    Main = workbook.add_worksheet('Main')

    keyCount = 1

    Main.set_column(im.size[0]+3, 4, 2)

    for y in range(im.size[1]):
        for x in range(im.size[0]):
            colourValue = '#%02x%02x%02x' % im.getpixel((x, y))[:3]
            if colourValue not in colours:
                colours.append(colourValue)
                Main.write(keyCount-1, 0, keyCount)
                Main.write(keyCount - 1, 1, '', workbook.add_format({"bg_color": colourValue}))
                Main.write(keyCount - 1, 2, colourValue)
                Main.write(y, x+4, len(colours), workbook.add_format({"bg_color": colourValue}))
                keyCount += 1
            else:
                Main.write(y, x+4, colours.index(colourValue)+1, workbook.add_format({"bg_color": colourValue}))

    workbook.close()
