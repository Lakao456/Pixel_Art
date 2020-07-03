import os
import pip
try:
    from PIL import Image
    import pandas as pd
except ImportError:
    pip.main(['install', 'pillow'])
    pip.main(['install', 'pandas'])
    from PIL import Image
    import pandas as pd

for image in os.listdir(os.getcwd()+'\\Input'):
    print(image)
    im = Image.open("Input/%s" % image)
    colours = []
    values = []

    for y in range(im.size[1]):
        values.append([])

        for x in range(im.size[0]):
            colourValue = '#%02x%02x%02x' % im.getpixel((x, y))[:3]
            if colourValue not in colours:
                colours.append(colourValue)
                values[y].append(len(colours))
            else:
                values[y].append(colours.index(colourValue)+1)


    art = pd.DataFrame(values)
    art.to_csv("Output_CSV\\PixelArt%s.csv" % image[:-4], index=False, header=False)
    key = pd.DataFrame(colours)
    key.index += 1
    key.to_csv("Output_CSV\\Key%s.csv" % image[:-4], header=False)
