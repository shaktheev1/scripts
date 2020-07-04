import pyexifinfo as p
import json
import os
from os import path
import csv

def get_meta():
    data1={}
    
    im = input("Enter folder path: ")
    image_directory = '/mnt/art/upload/'+im

    if path.exists(image_directory):
        imagelist=[]
        
        for dirName, subdirList, fileList in os.walk(image_directory):
                    for fname in fileList:
                        LowerCaseFileName = fname.lower()
                        if LowerCaseFileName.endswith(".jpg") or LowerCaseFileName.endswith(".eps") or LowerCaseFileName.endswith(".gif") or LowerCaseFileName.endswith(".png") or LowerCaseFileName.endswith(".tif") or LowerCaseFileName.endswith(".tiff") or LowerCaseFileName.endswith(".jpeg"):
                            imagelist.append(dirName+"/"+fname)
        csv_path = os.path.join(image_directory,'metadata.csv')
        with open(csv_path, 'w') as csvfile:
            csvwriter = csv.writer(csvfile)
            csvwriter.writerow(['Element number', 'Creator', 'Title', 'Description', 'Keywords', 'Source', 'Copyright notice', 'Rights usage terms'])
            
            for filename in imagelist:
                data = p.get_json(filename)
                element = os.path.basename(os.path.splitext(filename)[0])
                meta_list = data[0]
                # if 'XMP:Title' in meta_list:
                #     element = meta_list['XMP:Title']
                # else:
                #     element = ''
                if 'XMP:Creator' in meta_list:
                    creator = meta_list['XMP:Creator']
                else:
                    creator = ''    
                if 'XMP:Description' in meta_list:
                    title = meta_list['XMP:Description']
                    description = meta_list['XMP:Description']
                else:
                    title = ''
                    description = ''
                if 'IPTC:Keywords' in meta_list:
                    # keywords = ', '.join(meta_list['IPTC:Keywords'])
                    kwds=[]
                    kwds = meta_list['IPTC:Keywords']
                    keys = str(kwds).strip("[]")
                    keywords = str(keys).replace('\'','')
                else:
                    keywords = ''
                if 'XMP:Source' in meta_list:
                    source = meta_list['XMP:Source']
                else:
                    source = ''
                if 'EXIF:Copyright' in meta_list:
                    copy_right = meta_list['EXIF:Copyright']
                else:
                    copy_right = ''
                if 'XMP:UsageTerms' in meta_list:
                    usage = meta_list['XMP:UsageTerms']
                else:
                    usage = '' 
                
                csvwriter.writerow([element, creator, title, description, keywords, source, copy_right, usage])
    else:
        return "Folder with images does not exist inside {}!".format(image_directory)
    return "Metadata successfully created! The 'metadata.csv' is inside {}".format(image_directory)

result = get_meta()
print(result)
