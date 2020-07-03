import os
import pip

try:
    import excel2img
    from PIL import Image
except ImportError:
    pip.main(['install', 'pillow'])
    pip.main(['install', 'excel2img'])

os.chdir(os.getcwd() )

excelFiles = [item for item in os.listdir('Output')]+[item for item in os.listdir(os.getcwd()+'\\Output_coloured')]

print("\nSelect the list of files that you want to convert to PDF\n")
for i in range(len(excelFiles)):
    print("%d) %s" %(i, excelFiles[i]))

filesToConvert = [int(index) for index in input("Enter values separated by spaces:: ").split()]

for index in filesToConvert:
    outFolder = 'Output\\' if excelFiles[index] in os.listdir('Output') else 'Output_coloured\\'
    excel2img.export_img(outFolder+excelFiles[index], 'Output_Pdf\\'+excelFiles[index][:-5]+'.png', "Main", None)
    print(excelFiles[index], " Exported")

images = []

for image in os.listdir('Output_Pdf'):
    if image.lower().endswith('.png'):
        image = Image.open("Output_Pdf\\"+image)
        images.append(image.convert('RGB'))
        print(image, " Converted")

images[0].save('Output_Pdf\\' + input("Enter File Name:: ") + '.pdf', save_all=True, append_images=images[1:], resolution=100.0)

for image in os.listdir('Output_Pdf'):
    if image.lower().endswith('.png') or image.endswith('.jpg') or image.endswith('.jpeg'):
        os.remove("Output_Pdf\\"+image)
