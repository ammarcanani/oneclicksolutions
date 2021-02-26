import os
import datetime
import webbrowser
import shutil
import easygui



def createRootFolder():
    folderName = "Result/Files Duplicated at " + str(datetime.datetime.now().strftime("%d-%m-%y (Time %H-%M-%S)"))
    os.mkdir(folderName)
    return folderName

def createFolder(fileName):
    os.mkdir(fileName)
    return fileName

def openDonatePage():
    if("Donate Now.url" in os.listdir()):
        webbrowser.open("https://www.paypal.com/donate?hosted_button_id=3LYZG2R4962BN&source=url", new=0, autoraise=True)
        os.rename("Donate Now.url", "Support Us.url")

def namesOfFiles():
    file = open("NamesOfPersons.txt", "r")
    fileNames = []
    for line in file:
        fileNames.append(line.strip())
    file.close()
    return fileNames

def copyFiles(file, name, folderName):
    fileName = os.path.basename(file)
    extension = fileName[fileName.find("."):]
    filepath = folderName + "/"
    shutil.copy2(file, filepath + name + " (" + fileName + ")" + extension)

def main():
    # print(listFilesInFolder())
    # openDonatePage()
    filesToCopy = easygui.fileopenbox("Choose your file(s)", multiple=True)
    rootFolder = createRootFolder()
    namesOfFilesPerPerson = namesOfFiles()
    for person in namesOfFilesPerPerson:
        folderName = createFolder(rootFolder + "/" + person)
        for file in filesToCopy:
            copyFiles(file, person, folderName)
    openDonatePage()

if __name__ == '__main__':
    main()