##Arbortext to excel tables

from Tkinter import Tk
from tkFileDialog import askopenfilename

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file

plcodes = ["<pl0>","<pl1>","<pl2>","<pl3>","<pl4>","<pl5>","<pl6>","<pl7>","<pl8>","<pl9>"]
Allparts = []


class CreatePart(object):
    FigureNumber = "1"
    lastitem = ""
    def __init__(self, f):
        self.f = f
        self.splitit()

    def splitit(self):
        self.first = 0
        self.checkfignum()
        self.partno = self.eachpart("partno")
        self.cage = self.eachpart("cage")
        self.nomen = self.eachpart("nomen")
        self.units = self.eachpart("units")
        self.useoncode = self.eachpart("useoncode")
        self.smrcode = self.eachpart("smrcode")
        self.figindex = self.eachpart("figindex")
        if self.figindex == "":
            if self.first == 0:
                self.figindex = self.lastitem
        else:
            CreatePart.lastitem = self.figindex
        if self.units == "REF":
            self.figindex = ""

    def checkfignum(self):
        if "<figureno>" in self.f:
            start = self.f.find("<figureno>") + len("<figureno>")
            end = self.f.index("</figureno>", start)
            CreatePart.FigureNumber = self.f[start:end]
            self.first = 1
        self.figurenum = CreatePart.FigureNumber

    def eachpart(self, part):
        if part in self.f:
            start = self.f.find("<" + part+ ">") + len("<" + part+ ">")
            end = self.f.index("</" + part+ ">", start)
            return self.f[start:end].strip()
        else:
            return ""

with open(filename) as ipb:
    f = ""
    addit = "false"
    for i in ipb:
        if i.strip() != "":
            if i.strip()[:5] in plcodes:
                if "<useoncodelist>" in i:
                    addit = "false"
                else:
                    addit = "true"
                if f != "":
                    Allparts.append(CreatePart(f))
                
                f = i.strip()
            else:
                if "<useoncodelist>" in i:
                    addit = "false"
                
                if addit == "true":
                    f = f + i.strip() + " "
    Allparts.append(CreatePart(f))
    print len(Allparts)
    print Allparts[len(Allparts)-1].partno
    

with open("output.txt", "w") as out:
    for part in Allparts:
        line = part.figurenum + "\t" + part.figindex + "\t" + part.partno + "\t" + part.cage + "\t" + part.nomen + "\t" + part.units + "\t" + part.useoncode + "\t" + part.smrcode + "\n"
        out.write(line)
