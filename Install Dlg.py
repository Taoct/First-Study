# -*- coding: utf-8 -*-
# Author: Leo_Chen
# Time: 9/1/2018
import os
from os import walk
from Tkinter import *
import tkMessageBox
from ScrolledText import ScrolledText
import ConfigParser
import struct
import threading
import subprocess
from distutils.dir_util import copy_tree
import time
from time import strftime, localtime

# Define global Variable
g_StrVersion = '1.0.0.0'
g_strLogPath, PythonBitVersion = None, None  # removeNotSupportPackage(),removeNotSupportPackage()
PackageShowList, RefinePackageShowList, CheckVar = [], [], []
g_PackageRequirementList, RequirementList = [], []  # read ini(),removeNotSupportPackage()


def getDateTimeFormat():
    strDateTime = "[%s]" % (strftime("%Y/%m/%d %H:%M:%S", localtime()))
    return strDateTime


def printLog(strPrintLine):
    fileLog = open("Python_Package_Install_Tool.log", 'a')
    print strPrintLine
    fileLog.write("%s%s\n" % (getDateTimeFormat(), strPrintLine))
    fileLog.close()


def readINI(strINIPath):
    global g_strFailReasonTemp, g_nRecordSec, g_strEC, g_strdBUpperLimit, g_strdbLowerLimit, g_strplayingfile, RefinePackageShowList
    global g_nMicIndex, g_PackageRequirementList
    if not os.path.exists(strINIPath):
        g_strFailReasonTemp = "INI IS NOT EXIST"
        return False
    try:
        config = ConfigParser.ConfigParser()
        config.read(strINIPath)
        printLog("---------- INI ----------")
        for index in range(len(RefinePackageShowList)):
            if config.has_option('ConfigINI', RefinePackageShowList[index][0]):
                g_PackageRequirementList.append([RefinePackageShowList[index][0],
                                                 config.get('ConfigINI', RefinePackageShowList[index][0]).split(',')])
        for index in range(len(g_PackageRequirementList)):
            printLog(
                "[I][readINI] " + g_PackageRequirementList[index][0] + " = " + str(g_PackageRequirementList[index][1]))
        printLog("---------- INI ----------")
        return True
    # Error exception and raise error
    except ConfigParser.Error, err:
        g_strFailReasonTemp = ("Read INI Error: %s" % err)
        return False


# List remove is not actually delete
def removeNotSupportPackage():
    global PackageShowList, g_PIPFileName, RequirementList
    RemoveList = []
    if PythonBitVersion == 32:
        for i in range(len(PackageShowList)):
            # Check PackageShowList[i] (i=len(list)) exists amd64 characters,if exists,append it to removelist,then remove it
            if "amd64" in PackageShowList[i]:
                RemoveList.append(PackageShowList[i])
            if "pip-" in PackageShowList[i]:
                g_PIPFileName = PackageShowList[i]
                printLog("[I][Remove] Not Support PackageList" + g_PIPFileName)
                RemoveList.append(PackageShowList[i])
    elif PythonBitVersion == 64:
        for i in range(len(PackageShowList)):
            if "win32" in PackageShowList[i]:
                RemoveList.append(PackageShowList[i])
            if "pip-" in PackageShowList[i]:
                g_PIPFileName = PackageShowList[i]
                RemoveList.append(PackageShowList[i])
    # Remove action
    for item in range(len(RemoveList)):
        PackageShowList.remove(RemoveList[item])
    printLog("[I][removeNotSupportPackage] Remove Show List unsupported package success!")
    # -----------------------------------------------------------------------------------------------------------------------------#
    RemoveList = []
    if PythonBitVersion == 32:
        for i in range(len(RequirementList)):
            if "amd64" in RequirementList[i]:
                RemoveList.append(RequirementList[i])
    elif PythonBitVersion == 64:
        for i in range(len(RequirementList)):
            if "win32" in RequirementList[i]:
                RemoveList.append(RequirementList[i])

    # RemoveAction
    for item in range(len(RemoveList)):
        RequirementList.remove(RemoveList[item])
    printLog("[I][removeNotSupportPackage] Remove Requirement unsupported package success!")


def refinePackageShowList():
    global RefinePackageShowList
    # Making two-dimension array of[package name][package version] to make checkbox.
    RefinePackageShowList = [[None for x in range(2)] for y in range(len(PackageShowList))]
    for index in range(len(PackageShowList)):
        TempPackageName = PackageShowList[index]
        RefinePackageShowList[index][0] = TempPackageName.split('-')[0]
        RefinePackageShowList[index][1] = TempPackageName.split('-')[1]
    printLog("[I][RefinePackageShowList] CheckBox Refine name Successfully ")


def installPIP():
    PIPInstalledState = 0
    PackageInstallProcess = subprocess.Popen(['pip', 'install', './PythonPackageShowList' + g_PIPFileName], stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
    while True:
        time.sleep(1)
        error_line = PackageInstallProcess.stderr.readline()  # Process stdout message
        if error_line != "":
            # Find if package not support in this platform
            if "is not a supported" in error_line:
                PIPInstalledState = 2
                printLog("[W][Install Package]" + error_line.rstrip(
                    "\n"))  # Return a copy of the string S with trailing whitespace removed.
                break

        get_line = PackageInstallProcess.stderr.readline()
        if get_line != "":
            # Find if package already installed
            if "Requirement already satisfied" in get_line:
                PIPInstalledState = 1
                printLog("[W][Install Package] Requirement exists" + get_line.strip("\n"))
            else:
                PIPInstalledState = 0
                printLog("[W][Install Package]" + get_line.strip("\n"))
        else:
            break

    if PIPInstalledState == 1:
        PackageInstallDlg.printResult("[I] PIP already installed correctly!\n")
        printLog("[I][installPackage] PIP is already installed correctly!")
    elif PIPInstalledState == 2:
        PackageInstallDlg.printResult("[E] PIP Version ERROR!!!!\n")
        printLog("[E][installPackage] PIP Version Not Support!!")
    else:
        PackageInstallDlg.printResult("[I] PIP upgrade successfully!\n")
        printLog("[I][installPackage] PIP upgrade successfully!")

    PackageInstallDlg.inputControl("UnLock")
    PackageInstallDlg.changeRunningState("Ready")


def installpackage(P_ename):
    global strPyPath
    global PackageInstallProcess
    printLog('[I][installPackage] Installing ' + P_ename)
    bNotInRequirementDirectory = False

    if P_ename == "pyadb-0.1.4":
        strPyPath = g_strCurrentDirectory + "\\" + P_ename
        copy_tree(strPyPath, "C:\\Python27\\Lib\\site-packages\\pyadb")
    elif 'olefile' in P_ename:
        copy_tree(strPyPath, "C:\\Python27\\Lib\\site-packages\\pyadb")
        os.chdir("./PythonPackageRequirement/olefile-0.45.1")
        strolefilePath = g_strCurrentDirectory + '\\' + P_ename + '\\setup.py'
        PackageInstallProcess = subprocess.call(["python", strolefilePath, "install"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
        printLog("[I][startInstalling] olefile installed.")
        os.chdir("../../")
    else:
        # InstalledState: 0 = successfully installed ; 1 = Already installed; 2 = Not support in this platform.
        InstallState = 0
        PackageInstallProcess = subprocess.Popen(['pip', 'install', './PythonPackageRequirement/' + P_ename], stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
        # Read console error and message
        while True:
            errorline = PackageInstallProcess.stderr.readline()
            if errorline != "":
                # If Package is not in PythonPackageRequirement directory
                if 'looks like a filename, but the file does not exist' in errorline:
                    bNotInRequirementDirectory = True
                    break
                if "is not a supported" in errorline:
                    InstallState = 2
                    printLog("[W][installPackage] " + errorline.rstrip('\n'))
                    break
            line = PackageInstallProcess.stdout.readline()
            if line != "":
                # Find if package already installed
                if "Requirement already satisfied" in line:
                    InstallState = 1
                    printLog("[W][installPackage] " + line.rstrip('\n'))
                else:
                    InstallState = 0
                    printLog("[I][installPackage] " + line.rstrip('\n'))
            else:
                break
# ========================================================================================================================================
            # Not in PythonPackageRequirement Directory, Now try yo use PythonPackageShowList Directory
            if bNotInRequirementDirectory:
                InstallState = 0
                PackageInstallProcess = subprocess.Popen(['pip', 'install', './PythonPackageShowList/' + P_ename], stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
                # Read console error and message
                while True:
                    errorline = PackageInstallProcess.stderr.readline()
                    if errorline != "":
                        # Find if package not support in this platform
                        if "is not a supported" in errorline:
                            InstallState = 2
                            printLog("[W][installPackage] " + errorline.rstrip('\n'))
                            break
                    line = PackageInstallProcess.stdout.readline()
                    if line != "":
                        # Find if package already installed
                        if "Requirement already satisfied" in line:
                            InstallState = 1
                            printLog("[W][installPackage] " + line.rstrip('\n'))
                        else:
                            InstallState = 0
                            printLog("[I][installPackage] " + line.rstrip('\n'))
                    else:
                        break
        if InstallState == 1:
            PackageInstallDlg.printResult("[W] Already installed!\n")
            printLog("[W][installPackage] " + P_ename.split('-')[0] + " final state: Already installed!")
        elif InstallState == 2:
            PackageInstallDlg.printResult("[E] Wrong version!!\n")
            printLog("[E][installPackage] " + P_ename.split('-')[0] + " final state: Wrong version!")
        else:
            PackageInstallDlg.printResult("[I] Successfully Insalled!\n")
            printLog("[I][installPackage] " + P_ename.split('-')[0] + " final state: Successfully Insalled!")


class PackageInstallDlg(Frame):
    def __init__(self, master=None, width=0, height=0):
        Frame.__init__(self, master)
        self.createdlg()

    def createdlg(self):
        global nWindowWidth, nWindowHeight

        # window title
        self.master.title('PackageInstall_Tool %s' % (g_StrVersion))
        self.pack(fill=BOTH,
                  expand=1)  # expand = 1 can use fix ,and fill has x AND y AND both,3 parameters,both will expand to x and y ,when expand your GUI
        self.master.iconbitmap(r'Install.ico')
        self.master.protocol('WM_DELETE_WINDOW',
                             self.on_closing)  # define the DELETE_WINDOW ACTION is on_closing activity

        nWindowWidth, nWindowHeight = (self.winfo_screenwidth() * 3 / 4), (self.winfo_screenheight() * 3 / 4)
        nWindowsPosX = (self.winfo_screenwidth() / 2) - (nWindowWidth / 2)
        nWindowsPosY = (self.winfo_screenheight() / 2) - (nWindowHeight / 2)
        self.master.minsize(width=nWindowWidth, height=nWindowHeight)
        self.master.maxsize(width=nWindowWidth, height=nWindowHeight)
        self.master.geometry("+%d+%d" % (nWindowsPosX, nWindowsPosY))
        self.showReportBox()
        self.showExitBtn()
        self.showCheckFrame()
        self.showPackageCheckBox()
        self.showInstallBtn()
        self.showResetBtn()
        self.showStatus()
        self.checkPIP()

        # print self.winfo_screenwidth()
        # print nWindowWidth
        # print nWindowHeight
        # print nWindowsPosX
        # print nWindowsPosY

    def checkPIP(self):
        self.printResult("[I] Check PIP Version....\n")
        self.inputControl("Lock")
        self.changeRunningState("Running")
        threading.Thread(target=installPIP).start()  # Define Function installPIP

    # Result box generate
    def showReportBox(self):
        self.ResultBoxLabel = Label(self, text="Status Log:", anchor='w', foreground="#000080")
        self.ResultBoxLabel['font'] = ("Microsoft JhengHei", 16, "bold")
        self.ResultBoxLabel.place(x=nWindowWidth * 40 / 80, y=nWindowHeight * 2 / 80, width=nWindowWidth * 40 / 80,
                                  height=nWindowHeight * 4 / 80)
        self.ResultBox = ScrolledText(self, state=DISABLED, foreground="#000080")  # Load ScrolledText from dll
        self.ResultBox.place(x=nWindowWidth * 40 / 80, y=nWindowHeight * 6 / 80, width=nWindowWidth * 39 / 80,
                             height=nWindowHeight * 60 / 80)

    def showExitBtn(self):
        self.ExitBtn = Button(self, text="Exit", command=self.exitButtonDlg, foreground="#000080")
        self.ExitBtn.place(x=nWindowWidth * 68 / 80, y=nWindowHeight * 69 / 80, width=nWindowWidth * 10 / 80, height=nWindowHeight * 8 / 80)
        self.ExitBtn['font'] = ("Microsoft JhengHei", 18, "bold")

    def showResetBtn(self):
        self.retBtn = Button(self, text="Reset", command=self.resetButtonDlg, foreground="#000080")
        self.retBtn.place(x=nWindowWidth * 55 / 80, y=nWindowHeight * 69 / 80, width=nWindowWidth * 10 / 80, height=nWindowHeight * 8 / 80)
        self.retBtn['font'] = ("Microsoft JhengHei", 18, "bold")

    def resetButtonDlg(self):
        if tkMessageBox.askokcancel("Quit", "Do you want to reset and continue to start install !"):
            self.inputControl("UnLock")
            self.changeRunningState("Ready")


    def exitButtonDlg(self):
        if tkMessageBox.askokcancel("Quit", "Do you want to close the window!"):
            self.master.destroy()

    def on_closing(self):
        if tkMessageBox.askokcancel("Quit", "Do you want to close the windows?"):  # tkMessageBox.xxxxx(title, message, options)
            self.master.destroy()

    def showCheckFrame(self):
        self.CheckFrame = LabelFrame(self, text="-.-.- Package List -.-.-", labelanchor="n", foreground="orange")
        self.CheckFrame['font'] = ("Microsoft JhengHei", 16, "bold")
        self.CheckFrame.place(x=nWindowWidth * 2 / 80, y=nWindowHeight * 10 / 80, width=nWindowWidth * 37 / 80,
                              height=nWindowHeight * 65 / 80)

    # Package List Check Box generate
    def showPackageCheckBox(self):
        global CheckVar
        CheckVar = []
        self.ListCheckbox = []
        for PackageShowListIndex in range(len(PackageShowList)):
            self.var = IntVar()
            ##############################
            self.ListCheckbox.append(Checkbutton(self, text=str(
                RefinePackageShowList[PackageShowListIndex][0]) + " [ version: " + str(
                RefinePackageShowList[PackageShowListIndex][1]) + " ] ", variable=self.var, anchor="w"))
            CheckVar.append(self.var)
            self.ListCheckbox[PackageShowListIndex].place(x=nWindowWidth * 3 / 80, y=(nWindowHeight * 14 / 80) + (
                        PackageShowListIndex * nWindowHeight * 9 / 320), width=nWindowWidth * 20 / 80,
                                                          height=nWindowHeight * 2 / 80)
        self.printResult("List check box has Initialed\n")
        printLog("[I][showPackageCheckBox] Checkbox successfully generate")

    # Check and Uncheck and install button
    def showInstallBtn(self):
        self.checkAllBtn = Button(self, text="Select All", command=self.checkAll, foreground="#000080")
        self.checkAllBtn.place(x=nWindowWidth * 4 / 80, y=nWindowHeight * 65 / 80, width=nWindowWidth * 15 / 80,
                               height=nWindowHeight * 8 / 80)
        self.checkAllBtn['font'] = ("Microsoft JhengHei", 18, "bold")

        self.uncheckAllBtn = Button(self, text="Unselect All", command=self.uncheckAll, foreground="#000080")
        self.uncheckAllBtn.place(x=nWindowWidth * 22 / 80, y=nWindowHeight * 65 / 80, width=nWindowWidth * 15 / 80,
                                 height=nWindowHeight * 8 / 80)
        self.uncheckAllBtn['font'] = ("Microsoft JhengHei", 18, "bold")

        self.InstallBtn = Button(self, text="Start Install", command=self.installBtn, foreground="#af0000")
        self.InstallBtn.place(x=nWindowWidth * 40 / 80, y=nWindowHeight * 69 / 80, width=nWindowWidth * 12 / 80,
                              height=nWindowHeight * 8 / 80)
        self.InstallBtn['font'] = ("Microsoft JhengHei", 18, "bold")

    def checkAll(self):
        for i in self.ListCheckbox:
            i.select()

    def uncheckAll(self):
        for i in self.ListCheckbox:
            i.deselect()
        self.changeRunningState("UnSelect")

    def installBtn(self):
        self.changeRunningState("Running")
        self.printResult("Install is running...\n")
        self.inputControl("Lock")
        threading.Thread(target=self.startInstalling).start()

    # Status label text and place
    def showStatus(self):
        self.StatusText = Label(self, text="State: ", foreground="#000080")
        self.StatusText['font'] = ("Microsoft JhengHei", 26, "bold")
        self.StatusText.place(x=nWindowWidth * 2 / 80, y=nWindowHeight * 1 / 80, width=nWindowWidth * 8 / 80,
                              height=nWindowHeight * 8 / 80)

        self.StatusLabel = Label(self, text="Ready......  ", bg="green")
        self.StatusLabel['font'] = ("Microsoft JhengHei", 26, "bold")
        self.StatusLabel.place(x=nWindowWidth * 10 / 80, y=nWindowHeight * 1 / 80, width=nWindowWidth * 16 / 80,
                               height=nWindowHeight * 8 / 80)

    # Status Label change Function
    def changeRunningState(self, RunningorReady):
        if RunningorReady == "Running":
            self.printResult("Clicked installbtn,Ready Install...\n")
            self.StatusLabel.destroy()
            self.StatusLabel = Label(self, text="Running", bg="red")
            self.StatusLabel.place(x=nWindowWidth * 5 / 40, y=nWindowHeight * 1 / 80, width=nWindowWidth * 8 / 40)
            self.StatusLabel['font'] = ("Microsoft JhengHei", 26, "bold")
        if RunningorReady == "Ready":
            self.printResult("Getting on ready...\n")
            self.StatusLabel.destroy()
            self.StatusLabel = Label(self, text="Ready", bg="green")
            self.StatusLabel.place(x=nWindowWidth * 5 / 40, y=nWindowHeight * 1 / 80, width=nWindowWidth * 8 / 40)
            self.StatusLabel['font'] = ("Microsoft JhengHei", 26, "bold")
        if RunningorReady == "Finish":
            self.printResult("Process has finished...\n")
            self.StatusLabel.destroy()
            self.StatusLabel = Label(self, text="Finish", bg="blue")
            self.StatusLabel.place(x=nWindowWidth * 5 / 40, y=nWindowHeight * 1 / 80, width=nWindowWidth * 8 / 40)
            self.StatusLabel['font'] = ("Microsoft JhengHei", 26, "bold")
        if RunningorReady == "UnSelect":
            self.printResult("You Clicked Unselect command...\n")
            self.StatusLabel.destroy()
            self.StatusLabel = Label(self, text="UnSelect", bg="yellow")
            self.StatusLabel.place(x=nWindowWidth * 5 / 40, y=nWindowHeight * 1 / 80, width=nWindowWidth * 8 / 40)
            self.StatusLabel['font'] = ("Microsoft JhengHei", 26, "bold")

    def inputControl(self, L_state):

        if L_state == "Lock":
            for item in self.ListCheckbox:
                item["state"] = "disabled"
            self.checkAllBtn['state'] = 'disabled'
            self.uncheckAllBtn['state'] = 'disabled'
            self.InstallBtn['state'] = 'disabled'
        if L_state == "UnLock":
            for item in self.ListCheckbox:
                item["state"] = "normal"
            self.checkAllBtn['state'] = 'normal'
            self.uncheckAllBtn['state'] = 'normal'
            self.InstallBtn['state'] = 'normal'

    # Result box status input
    def printResult(self, str_Info):
        self.ResultBox.configure(state='normal')
        self.ResultBox.insert(END, str_Info)
        self.ResultBox.configure(state='disabled')

    def startInstalling(self):
        for i in range(len(PackageShowList)):
            if CheckVar[i].get() == 1:
                PackageInstallDlg.printResult("[I] Searching " + PackageShowList[i] + " requirement...\n")
                for index in range(len(g_PackageRequirementList)):
                    # requirement in show list
                    if g_PackageRequirementList[index][0] in PackageShowList[i]:
                        for innerindex in range(len(g_PackageRequirementList[index][1])):
                            for j in range(len(RequirementList)):
                                if g_PackageRequirementList[index][1][innerindex] in RequirementList[j]:
                                    PackageInstallDlg.printResult("[I] Installing: " + RequirementList[j] + "\n")
                                    installpackage(RequirementList[j])
                PackageInstallDlg.printResult("[I] Installing: " + PackageShowList[i] + "\n")
                installpackage(PackageShowList[i])
                print "end"

        # State labe change
        self.changeRunningState("Finish")
        self.printResult("╔═══════════════╗\n")
        self.printResult("║         F i n i s h !        ║\n")
        self.printResult("╚═══════════════╝\n")


if __name__ == '__main__':
    # python bit version
    PythonBitVersion = struct.calcsize("P") * 8

    # Get Current Working Directory
    g_strCurrentDirectory = os.path.dirname(os.path.abspath(__file__)) + "\\PythonPackageShowList"

    for (dirpath, dirnames, filenames) in walk(g_strCurrentDirectory):
        PackageShowList.extend(filenames)
        break
    for (dirpath, dirnames, filenames) in walk(g_strCurrentDirectory):
        PackageShowList.extend(dirnames)
        break

    g_strCurrentDirectory = os.path.dirname(os.path.abspath(__file__)) + "\\PythonPackageRequirement"
    for (dirpath, dirnames, filenames) in walk(g_strCurrentDirectory):
        RequirementList.extend(filenames)
        break
    for (dirpath, dirnames, filenames) in walk(g_strCurrentDirectory):
        RequirementList.extend(dirnames)
        break

    removeNotSupportPackage()
    refinePackageShowList()

    strINIPath = os.path.dirname(os.path.abspath(__file__)) + "\\Python_Package_Install_Tool.ini"
    readINI(strINIPath)

    cTk = Tk()
    PackageInstallDlg = PackageInstallDlg(master=cTk)
    PackageInstallDlg.mainloop()
