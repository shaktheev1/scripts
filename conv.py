import os, sys
from PIL import Image
from fnmatch import fnmatch

# for infile in os.listdir("."):
#     #print "file : " + infile
#     if infile[-3:] == "png" or infile[-3:] == "tif" :
#        # print "is tif or bmp"
#        outfile = infile[:-3] + "jpg"
#        im = Image.open(infile)
#        #print "new filename : " + outfile
#        out = im.convert("RGB")
#        out.save(outfile, "JPEG", quality=90)


images='images'

root=input("Enter folder name:")
for infile in os.listdir(root):
    if infile.endswith('.png') or infile.endswith('.tif') or infile.endswith('.jpg') or infile.endswith('.eps'):
        im = Image.open(os.path.join(root,infile))
        if not(os.path.exists(os.path.join(root,images))):
            os.mkdir(os.path.join(root,images))
            os.chmod(os.path.join(root,images), 0o02775)
            os.system("chgrp -R space "+os.path.join(root,images))
        outfile = os.path.join(root,images,os.path.splitext(infile)[0]+".jpg")
        out = im.convert("RGB")
        out.save(outfile, "JPEG", quality=90)



