import xlrd

file_location = "G:\python\\New folder\\test6.xls"
workbook = xlrd.open_workbook(file_location)

sheet = workbook.sheet_by_name("overall")
nForFrame = sheet.nrows - 1

#print(sheet.cell_value(2,4))
x1 = []
y1 = []
z1 = []
x2 = []
y2 = []
z2 = []
for row in range(1, sheet.nrows):
	x1_1 = sheet.cell_value(row,0)
	y1_1 = sheet.cell_value(row,1)
	z1_1 = sheet.cell_value(row,2)
	x2_1 = sheet.cell_value(row,3)
	y2_1 = sheet.cell_value(row,4)
	z2_1 = sheet.cell_value(row,5)
	x1.append(x1_1)
	y1.append(y1_1)
	z1.append(z1_1)	
	x2.append(x2_1)
	y2.append(y2_1)
	z2.append(z2_1)

'''sheet = workbook.sheet_by_name("Model Details")
nForStory = sheet.nrows - 1

StoryName = []
StoryHeight = []
StoryElevation = []
MasterStory = []
SimilarStory = []
SpliceAbove = []
SpliceHeight = []
NoOfStory = int()
for i in range(1, sheet.nrows):
	s1 = sheet.cell_value(i,0)
	s2 = sheet.cell_value(i,1)
	s3 = sheet.cell_value(i,2)
	s4 = sheet.cell_value(i,3)
	s5 = sheet.cell_value(i,4)
	s6 = sheet.cell_value(i,5)
	s7 = sheet.cell_value(i,6)
	StoryName.append(s1)
	StoryHeight.append(s2)
	StoryElevation.append(s3)
	MasterStory.append(bool(s4))
	SimilarStory.append(s5)
	SpliceAbove.append(bool(s6))
	SpliceHeight.append(s7)
print(nForStory)
print(nForFrame)
print(StoryElevation)
print(MasterStory)
print(SimilarStory)
print(SpliceAbove)
print(SpliceHeight)
print(type(StoryName[1]))
print(type(StoryHeight[1]))'''


import comtypes
import os
import sys
import comtypes.client
import numpy

APIPath = "G:\etabs\API"
if not os.path.exists(APIPath):
    os.makedirs(APIPath)
model_path = APIPath + os.sep + "Sample 0001.edb"


pro_path = "C:\Program Files\Computers and Structures\ETABS 2016\ETABS.exe"

helper = comtypes.client.CreateObject("ETABS2016.Helper")
helper = helper.QueryInterface(comtypes.gen.ETABS2016.cHelper)
myETABSObject=helper.CreateObject(pro_path)
myETABSObject.ApplicationStart()


#create SapModel object
SapModel = myETABSObject.SapModel

#initialize model with units
Units = 6
SapModel.InitializeNewModel(Units)


#create new blank model
ret = SapModel.File.NewBlank()

#switch to k-ft units
#kip_ft_F = 3
#ret = SapModel.SetPresentUnits(kip_ft_F)

#define material property
MATERIAL_CONCRETE = 2
ret = SapModel.PropMaterial.SetMaterial('CONC', MATERIAL_CONCRETE)


#assign isotropic mechanical properties to material
ret = SapModel.PropMaterial.SetMPIsotropic('CONC', 27, 0.2, 0.0000055)

#define rectangular frame section property
ret = SapModel.PropFrame.SetRectangle('R1', 'CONC', 0.4, 0.4)

#define frame section property modifiers
ModValue = [1000, 0, 0, 1, 1, 1, 1, 1]
ret = SapModel.PropFrame.SetModifiers('R1', ModValue)

#add frame object by coordinates
FrameName = ' '

for i in range(nForFrame):
	[FrameName , ret] = SapModel.FrameObj.AddByCoord(x1[i] , y1[i] , z1[i] , x2[i] , y2[i] , z2[i] , FrameName , 'R1' , 'Global')

'''#set story limits
MyStory =' ' 
MyStory2 = ' '

for m in range(int(nForStory)):
    [MyStory, ret] = SapModel.Story.SetStories('Story' + str([m]), StoryElevation[m], StoryHeight[m], False , 'none' , False , 0 )
    '''
    #[MyStory2 , ret] = SapModel.Story.GetStories(NoOfStory, StoryName[i], StoryElevation[i], StoryHeight[i], MasterStory[i], SimilarStory[i], SpliceAbove[i], SpliceHeight[i]) 
'''#assign point object restraint at base
PointName1 = ' '
PointName2 = ' '
Restraint = [True, True, True, True, False, False]
[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName, PointName1, PointName2)
ret = SapModel.PointObj.SetRestraint(PointName1, Restraint)

#assign point object restraint at top
Restraint = [True, True, False, False, False, False]
[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName, PointName1, PointName2)
ret = SapModel.PointObj.SetRestraint(PointName2, Restraint)'''

#refresh view, update (initialize) zoom
ret = SapModel.View.RefreshView(0, False)

'''#add load patterns
LTYPE_OTHER = 8
ret = SapModel.LoadPatterns.Add('1', LTYPE_OTHER, 1, True)
ret = SapModel.LoadPatterns.Add('2', LTYPE_OTHER, 0, True)
ret = SapModel.LoadPatterns.Add('3', LTYPE_OTHER, 0, True)
ret = SapModel.LoadPatterns.Add('4', LTYPE_OTHER, 0, True)
ret = SapModel.LoadPatterns.Add('5', LTYPE_OTHER, 0, True)
ret = SapModel.LoadPatterns.Add('6', LTYPE_OTHER, 0, True)
ret = SapModel.LoadPatterns.Add('7', LTYPE_OTHER, 0, True)

#assign loading for load pattern 2
[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName3, PointName1, PointName2)
PointLoadValue = [0,0,-10,0,0,0]
ret = SapModel.PointObj.SetLoadForce(PointName1, '2', PointLoadValue)
ret = SapModel.FrameObj.SetLoadDistributed(FrameName3, '2', 1, 10, 0, 1, 1.8, 1.8)

#assign loading for load pattern 3
[PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints(FrameName3, PointName1, PointName2)
PointLoadValue = [0,0,-17.2,0,-54.4,0]
ret = SapModel.PointObj.SetLoadForce(PointName2, '3', PointLoadValue)

#assign loading for load pattern 4
ret = SapModel.FrameObj.SetLoadDistributed(FrameName2, '4', 1, 11, 0, 1, 2, 2)

#assign loading for load pattern 5
ret = SapModel.FrameObj.SetLoadDistributed(FrameName1, '5', 1, 2, 0, 1, 2, 2, 'Local')
ret = SapModel.FrameObj.SetLoadDistributed(FrameName2, '5', 1, 2, 0, 1, -2, -2, 'Local')

#assign loading for load pattern 6
ret = SapModel.FrameObj.SetLoadDistributed(FrameName1, '6', 1, 2, 0, 1, 0.9984, 0.3744, 'Local')
ret = SapModel.FrameObj.SetLoadDistributed(FrameName2, '6', 1, 2, 0, 1, -0.3744, 0, 'Local')

#assign loading for load pattern 7
ret = SapModel.FrameObj.SetLoadPoint(FrameName2, '7', 1, 2, 0.5, -15, 'Local')'''

'''#switch to k-in units
kip_in_F = 2
ret = SapModel.SetPresentUnits(kip_in_F)
'''
#save model
ret = SapModel.File.Save(model_path)

#run model (this will create the analysis model)
ret = SapModel.Analyze.RunAnalysis()

#close the program
ret = myETABSObject.ApplicationExit(False)
SapModel = None
myETABSObject = None
