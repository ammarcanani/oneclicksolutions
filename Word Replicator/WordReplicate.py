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
    fileName = fileName[: fileName.find(".")]
    os.mkdir(fileName)
    return fileName

def openDonatePage():
    if("Donate Now.url" in os.listdir()):
        webbrowser.open("https://www.paypal.com/donate?hosted_button_id=3LYZG2R4962BN&source=url", new=0, autoraise=True)
        os.rename("Donate Now.url", "Support Us.url")

def namesOfFiles():
    file = open("NamesOfFiles.txt", "r")
    fileNames = []
    for line in file:
        fileNames.append(line.strip())
    file.close()
    return fileNames

def copyFiles(fileName, namesOfFiles, folderName):
    filePathToCopy = fileName
    extension = fileName[fileName.find("."):]
    nameOfFile = os.path.basename(fileName)[:fileName.find(".")]
    for name in namesOfFiles:
        # print(filePathToCopy, folderName + "/" + name + " (" + nameOfFile + ")" + extension)
        shutil.copy2(filePathToCopy, folderName + "/" + name + " (" + nameOfFile + ")" + extension)

def main():
    # print(listFilesInFolder())
    # openDonatePage()
    print("Please select your file(s)...")
    filesToCopy = easygui.fileopenbox("Choose your file(s)", multiple=True, filetypes=[["*.doc", "*.docx", "Word Files"]], default="*.doc*")
    # print(filesToCopy)
    rootFolder = createRootFolder()
    print("Processing...")
    for file in filesToCopy:
        fileName = os.path.basename(file)
        folderName = createFolder(rootFolder + "/" + fileName)
        copyFiles(file, namesOfFiles(), folderName)
    print("Files Replicated...")
    print("Don't forget to support us!")
    openDonatePage()

if __name__ == '__main__':
    print("Thank you for using File Replicator")
    main()
