import xlsxwriter
import json
import modelList as m
import refFile as r
import os
from tkinter import *
from tkinter import simpledialog
from tkinter import messagebox

sensorcluandattribute = {}
refsecuritysensor = {}
deviceCluAttributes = {}
outputPass = {}
outputFail = {}
outputPNR = {}

window = Tk()  # UI for userinput - Mac ID
window.withdraw()

macId = simpledialog.askstring(title="ZigBeeServiceValidationTool",
                               prompt="Enter The Device MAC ID:")
deviceCurrentFilePath = os.getcwd()
deviceFilePath = os.path.join(deviceCurrentFilePath, "InputFile", macId)

try:
    with open(deviceFilePath, 'r') as f:
        deviceCluAtt = json.load(f)
except OSError as e:
    messagebox.showinfo("Error", f"Mac ID {macId} Not Found")
    exit()

deviceCluAttributes = deviceCluAtt

deviceRefFile = r.refCluAttribute  # get reference clusters and attribute from refFile

modelnumber = deviceCluAtt['metadatas']['discoveredDetails']['value']['model']

legacyDevice = deviceCluAtt['metadatas']  # check if the device is legacy device
if legacyDevice.__contains__("legacyDevType"):
    messagebox.showinfo("Error", "It is a Legacy Device")
    exit()
else:
    deviceTypeList = m.deviceType  # get all the device types from modelList


    # To identify the deviceType (eg: doorlock, DWS)
    def get_key(modelnumber):
        for key in deviceTypeList.keys():
            modellist = deviceTypeList[key]
            if modellist.__contains__(modelnumber):
                return str(key)

        return ""   #messagebox.showinfo("Error", f"Device Model '{modelnumber}' Not Found")


    deviceType = (get_key(modelnumber))

    if not deviceType.__len__() > 0:
        types = deviceTypeList.keys()
        master = Tk()
        variable = StringVar(master)
        variable.set(0)  # default value
        w = OptionMenu(master, variable, *types)


    # To identify the reference clus and attribute based on the model number
    def get_ref(deviceType):
        for key in deviceRefFile.keys():
            if key == deviceType:
                return deviceRefFile[key]


    deviceReferenceFile = get_ref(deviceType)

    refsecuritysensor = deviceReferenceFile

    for cluster in deviceCluAttributes['metadatas']['discoveredDetails']['value']['endpoints']:
        servercluster = cluster['serverClusterInfos']
        clientcluster = cluster['clientClusterInfos']
        for sercluster in servercluster:
            serclustiddec = hex(sercluster['id'])
            serclusattribute = sercluster['attributeIds']
            serattributedeclist = []
            for serattributes in serclusattribute:
                serattributedeclist.append(hex(serattributes))
            sensorcluandattribute.update({serclustiddec: serattributedeclist})
        for clicluster in clientcluster:
            cliclustiddec = hex(clicluster['id'])
            cliclusattribute = clicluster['attributeIds']
            cliclusandattributelist = []
            for cliattributes in cliclusattribute:
                cliclusandattributelist.append(hex(cliattributes))
            if sensorcluandattribute.__contains__(
                    cliclustiddec):  # to handle overriding if ser and cli clusters are same
                cliclusandattributelist.append(sensorcluandattribute.get(cliclustiddec))
            sensorcluandattribute.update({cliclustiddec: cliclusandattributelist})

        for key in sensorcluandattribute:
            if refsecuritysensor.__contains__(key):
                passattribute = []
                pnrattribute = []
                for senatt in sensorcluandattribute[key]:
                    if refsecuritysensor[key].__contains__(senatt):
                        passattribute.append(senatt)
                        outputPass.update({key: passattribute})
                    else:
                        pnrattribute.append(senatt)
                        outputPNR.update({key: pnrattribute})
            else:
                if sensorcluandattribute[key].__len__() > 0:
                    pnroneattribute = []
                    for senatt in sensorcluandattribute[key]:
                        pnroneattribute.append(senatt)
                        outputPNR.update({key: pnroneattribute})
                else:
                    outputPNR[key] = []
        for key in refsecuritysensor:
            if not sensorcluandattribute.__contains__(key):
                if refsecuritysensor[key].__len__() > 0:
                    failattribute = []
                    for refatt in refsecuritysensor[key]:
                        failattribute.append(refatt)
                        outputFail.update({key: failattribute})
                else:
                    outputFail[key] = []
            else:
                failoneatribute = []
                for refatt in refsecuritysensor[key]:
                    if not sensorcluandattribute[key].__contains__(refatt):
                        failoneatribute.append(refatt)
                        outputFail.update({key: failoneatribute})

    result = ""
    if outputFail.__len__() == 0:
        result = "PASS"
    else:
        result = "FAIL"

    if result == "PASS":
        messagebox.showinfo("Result", "PASS - Find more details in Result Directory")
    elif result == "FAIL":
        messagebox.showinfo("Result", "FAIL - Find missing Cluster or Attribute details in Result Directory")

    deviceFilePathresults = os.path.join(deviceCurrentFilePath, "Results", deviceType + '.xlsx')

    wb = xlsxwriter.Workbook(deviceFilePathresults)

    sheet1 = wb.add_worksheet('PASS')
    sheet2 = wb.add_worksheet('ABSENT')
    sheet3 = wb.add_worksheet('PNR')

    green = wb.add_format()
    green.set_pattern(1)
    green.set_bg_color('green')
    green.set_bold({'bold': True})

    red = wb.add_format()
    red.set_pattern(1)
    red.set_bg_color('red')
    red.set_bold({'bold': True})

    bold = wb.add_format({'bold': True})

    sheet1.write(0, 2, "Status", bold)
    sheet1.write(0, 0, "Cluster", bold)
    sheet1.write(0, 1, "Attribute", bold)
    sheet2.write(0, 0, "Cluster", bold)
    sheet2.write(0, 1, "Attribute", bold)
    sheet3.write(0, 0, "Cluster", bold)
    sheet3.write(0, 1, "Attribute", bold)

    if result == "FAIL":
        sheet1.write('D1', result, red)
    else:
        sheet1.write('D1', result, green)

    passrow = 0
    passcol = 0
    failrow = 0
    failcol = 0
    pnrrow = 0
    pnrcol = 0

    for clu in outputPass.keys():
        passrow += 1
        sheet1.write(passrow, passcol, clu)
        for att in outputPass[clu]:
            sheet1.write(passrow, passcol + 1, att)
            passrow += 1

    for clu in outputFail.keys():
        failrow += 1
        sheet2.write(failrow, failcol, clu)
        for att in outputFail[clu]:
            sheet2.write(failrow, failcol + 1, att)
            failrow += 1

    for clu in outputPNR.keys():
        pnrrow += 1
        sheet3.write(pnrrow, pnrcol, clu)
        for att in outputPNR[clu]:
            sheet3.write(pnrrow, pnrcol + 1, att)
            pnrrow += 1

    wb.close()
