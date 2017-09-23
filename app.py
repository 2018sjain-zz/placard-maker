from flask import Flask, request, render_template, send_from_directory
import sys
import zipfile
import os
import shutil

def generate(conference, place, names):
	conf = conference
	room = place
	vals = names

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

app = Flask(__name__)

@app.route("/")
def index():
	return render_template("index.html")

@app.route('/', methods=['POST'])
def my_form_post():
    conference = request.form['conference']
    room = request.form['room']
    people = request.form['people'].split(';')
    people = [person.strip() for person in people]

    generate(conference, room, people)

    people = ", ".join(x for x in people)
    return render_template("output.html", conference=conference, room=room, people=people)

@app.route('/<path:filename>', methods=['POST', 'GET'])
def download(filename):
	return send_from_directory('files/tmp', filename, as_attachment=True)

if __name__ == "__main__":
	app.run(debug = True)