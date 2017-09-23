import zipfile
import os
import shutil

conf = input("Conference Name?: ")
room = input("Committee Name?: ")
vals = open(input("Name of '.txt' file with Position Names\nexample - 'who.txt': ")).readlines()

newpath = ''+room #INPUT THE PATH OF THE DIRECTORY YOU ARE IN
os.makedirs(newpath)
for x in vals:
    name = x.rstrip()
    replacePerson = {"POSITION" : name}
    replaceRoom =   {"COMMITTEE" : room}
    replaceConf =   {"CONFERENCE": conf}
    templateDocx = zipfile.ZipFile('files/basic.docx') #INPUT THE PATH OF THE DIRECTORY YOU ARE IN
    newDocx = zipfile.ZipFile(room+'/'+name+room+".docx", "a")

    with open(templateDocx.extract("word/document.xml", "")) as tempXmlFile: #INPUT THE PATH OF THE DIRECTORY YOU ARE IN
        tempXmlStr = tempXmlFile.read()

    for key in replaceConf.keys():
        tempXmlStr = tempXmlStr.replace(str(key), str(replaceConf.get(key)))

    for key in replacePerson.keys():
        tempXmlStr = tempXmlStr.replace(str(key), str(replacePerson.get(key)))

    for key in replaceRoom.keys():
        tempXmlStr = tempXmlStr.replace(str(key), str(replaceRoom.get(key)))

    with open("temp.xml", "w+") as tempXmlFile: #INPUT THE PATH OF THE DIRECTORY YOU ARE IN
        tempXmlFile.write(tempXmlStr)

    for file in templateDocx.filelist:
        if not file.filename == "word/document.xml":
            newDocx.writestr(file.filename, templateDocx.read(file))

    newDocx.write("temp.xml", "word/document.xml") #INPUT THE PATH OF THE DIRECTORY YOU ARE IN

    templateDocx.close()
    newDocx.close()

    os.remove('temp.xml')
    os.remove('word/document.xml')
    os.rmdir('word')
if(os.path.isdir("files/tmp")):
    shutil.rmtree("files/tmp")
shutil.make_archive("files/tmp/"+conf+"_"+room, 'zip', room)
shutil.rmtree(room)