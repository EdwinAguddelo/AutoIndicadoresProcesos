import os.path

class pathFiles():
    def __init__(self,resourcesPath):
        self.resourcesPath = resourcesPath
        self.resourcesXlsx=[]
        for file in os.listdir(self.resourcesPath):
            if file.endswith(".xlsx"):
                self.resourcesXlsx.append(os.path.join(self.resourcesPath, file))

        self.OC_FilePath = self.resourcesXlsx[1]
        self.DOD_FilePath = self.resourcesXlsx[0]
