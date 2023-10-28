import zipfile
import os
from PIL import Image

# Traverse all files under a specified path
def findAllFile(base):
    for root, ds, fs in os.walk(base):
        for f in fs:
            if f.endswith('.jpeg') or f.endswith('.png'):
                fullname = os.path.join(root, f)
                yield fullname

# The path of the docx file to be compressed
zip_path = 'C:\\Users\\XX\\Desktop\\new1.docx'

# Used to save the decompressed docx file
save_path = 'C:\\Users\\XX\\Desktop\\nn\\'

# Decompress the docx file
file = zipfile.ZipFile(zip_path)
file.extractall(save_path)
file.close()

# The path of pictures in the decompressed docx file
picture_path = save_path+'word\\media\\'

# Compress pictures
# There is using pillow lib
for i in findAllFile(picture_path):
    # todo:
    img_path = i
    im = Image.open(i)
    # quality means compress rate, between 1-95?
    im.save(i, quality=40)

# The path to the generated new docx file
new_file_path = 'C:\\Users\\XX\\Desktop\\new1_com.docx'

# using append mode to open the generated new docx file
# specifically, don't compress, just pack it.
fz = zipfile.ZipFile(new_file_path, 'a')
# Traverse all files need to be compressed.
for root, ds, fs in os.walk(save_path):
    for f in fs:
            fullname = os.path.join(root, f)
            fpath = fullname.replace(save_path, '')
            fz.write(fullname, fpath)

fz.close()

