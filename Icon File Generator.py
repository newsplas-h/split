import base64
#Generates a "icon.py" file in the directory for use with executables
#base64 required import for all files using this


#put the blank file here for encoding purposes, best with x64
with open('Split128.png', 'rb') as image_file:
    encodeString = base64.b64encode(image_file.read())
image_file.close()
write_data=("img = %s" % encodeString)
f=open("icon.py","w+")
f.write(write_data)
f.close()
