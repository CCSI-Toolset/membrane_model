###############################################################################
#           Automated Test of CCSI Hollow Fiber Gas Permeation Membranne Model Simulator
#           Runs on Aspen Customer Modeler 7.3
#           Uses Python 2.7, pywinauto 0.4.3
#           Gregory Pope, LLNL
#           May 8, 2013
#           For Aceptance Testing Hollow Fiber Gas Permeation Membranne Model v0.9.c1
#           V1.0
###############################################################################
#
import pywinauto
import time

def ChangeMouse(script):
     #Assumes Ghost mouse is visable on desktop and scripts are in My Documents folder
     #Change(or load intital) mouse scripts into Ghost Mouse
     #Allows Ghost Mouse to change scripts during automated test run to create numerous mouse behaviors 
     print('Change Mouse Scripts')
     w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'GhostMouse 3.2', class_name='AutoIt v3 GUI')[0])
     window = pwa_app.window_(handle=w_handle)
     window.SetFocus()
     
     window.MenuItem(u'&File').Click()
     window.MenuItem(u'&File->&Open').Click()
     w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Open file...', class_name='#32770')[0])
     window = pwa_app.window_(handle=w_handle)
     window.SetFocus()
     ctrl = window['ComboBox']
     ctrl.SetFocus()
     ctrl.ClickInput()
     ctrl.TypeKeys(script)
     ctrl.TypeKeys('{ENTER}')
     return

def SetMouseSpeed(speed):
     #Assumes Ghost mouse is visable on desktop
     #Change(or load intital) mouse speed setting into Ghost Mouse (speed is non-zero integer -10 to 10, 10 = 10 times faster, -10 1/10th speed)
     #Allows Ghost Mouse to change speeds during automated test run to create numerous mouse behaviors 
     w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'GhostMouse 3.2', class_name='AutoIt v3 GUI')[0])
     window = pwa_app.window_(handle=w_handle)
     window.SetFocus()
     window.MenuItem(u'&Options').Click()
     window.MenuItem(u'&Options->&Playback').Click()
     window.MenuItem(u'&Options->&Playback->Spee&d').Click()
     w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Speed setting', class_name='AutoIt v3 GUI')[0])
     window = pwa_app.window_(handle=w_handle)
     window.SetFocus()
     ctrl = window['Trackbar']
     ctrl.SetFocus()
     if (speed == 10): ctrl.TypeKeys('{END}')
     elif (speed == -10): ctrl.TypeKeys('{HOME}')
     elif (speed == 1): ctrl.TypeKeys('{END}'+'{LEFT 9}')
     elif (speed == 2): ctrl.TypeKeys('{END}'+'{LEFT 8}')
     elif (speed == 3): ctrl.TypeKeys('{END}'+'{LEFT 7}')
     elif (speed == 4): ctrl.TypeKeys('{END}'+'{LEFT 6}')
     elif (speed == 5): ctrl.TypeKeys('{END}'+'{LEFT 5}')
     elif (speed == 6): ctrl.TypeKeys('{END}'+'{LEFT 4}')
     elif (speed == 7): ctrl.TypeKeys('{END}'+'{LEFT 3}')
     elif (speed == 8): ctrl.TypeKeys('{END}'+'{LEFT 2}')
     elif (speed == 9): ctrl.TypeKeys('{END}'+'{LEFT 1}')
     elif (speed == -9): ctrl.TypeKeys('{HOME}'+'{RIGHT 1}')
     elif (speed == -8): ctrl.TypeKeys('{HOME}'+'{RIGHT 2}')
     elif (speed == -7): ctrl.TypeKeys('{HOME}'+'{RIGHT 3}')
     elif (speed == -6): ctrl.TypeKeys('{HOME}'+'{RIGHT 4}')
     elif (speed == -5): ctrl.TypeKeys('{HOME}'+'{RIGHT 5}')
     elif (speed == -4): ctrl.TypeKeys('{HOME}'+'{RIGHT 6}')
     elif (speed == -3): ctrl.TypeKeys('{HOME}'+'{RIGHT 7}')
     elif (speed == -2): ctrl.TypeKeys('{HOME}'+'{RIGHT 8}')
     else: print ('speed must be non-zero integer -10 to 10')
     ctrl = window['Ok']
     ctrl.Click()
     return


#File needed to do simualtion in Aspen Custom Modeler/AM_Untitled
Needed_file='HFGP.acmf'
#assure we do not type too fast over remote set up
keywait =.33

pwa_app = pywinauto.application.Application()

# Assure Ghost Mouse help application in tray using file drag_moving_bed.rms
raw_input('Assure Ghost Mouse app. on screen, hit enter')


# Start Application
pwa_app.start_("C:\Program Files (x86)\AspenTech\AMSystem V7.3\Bin\AspenModeler.exe")
print ('Started Application')


#Dismiss Registration if present
print ('Dismiss Registration')
if (pywinauto.timings.WaitUntilPasses(5,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Aspen Technology Product Registration', class_name='#32770')[0])):
     w_handle = pywinauto.findwindows.find_windows(title=u'Aspen Technology Product Registration', class_name='#32770')[0]
     window = pwa_app.window_(handle=w_handle)
     window.SetFocus()
     ctrl = window['Button2']
     ctrl.Click()


#Select and Configure Physical Properties
print ('Select Physical Properties')
w_handle = pywinauto.timings.WaitUntilPasses(10,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.Maximize()
window.TypeKeys('%t')
window.TypeKeys('g')

w_handle = pywinauto.timings.WaitUntilPasses(8,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Physical Properties Configuration', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['RadioButton']
ctrl.Click()
ctrl = window['Button2']
ctrl.Click()

print ('Select Chemicals')

w_handle = pywinauto.timings.WaitUntilPasses(10,0.5,lambda:pywinauto.findwindows.find_windows(title=u'PropsPlus.aprbkp - Aspen Properties V7.3 - aspenONE - [Components Specifications - Data Browser]')[0])
window = pwa_app.window_(handle=w_handle)
window.edit.TypeKeys('CO2\r') # for carbon dioxide
window.edit.TypeKeys('H2O\r') # for water
window.edit.TypeKeys('N2\r')  # for nitrogen
window.edit.TypeKeys('O2\r')  # for oxygen

w_handle = pywinauto.findwindows.find_windows(title=u'PropsPlus.aprbkp - Aspen Properties V7.3 - aspenONE - [Components Specifications - Data Browser]')[0]
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#Tools->Next
window.TypeKeys('%t')
window.TypeKeys('n')

# Select Property Method
print ('Select Property Method')
ctrl = window['AfxOleControl9023']
ctrl.SetFocus()
ctrl = window['ComboBox14']
ctrl.Click()
ctrl.Select('PR-BM')
ctrl.Click()

#Next Step
#Tools->Next
window.TypeKeys('%t')
window.TypeKeys('n')


w_handle = pywinauto.timings.WaitUntilPasses(5,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Required Properties Input Complete', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

w_handle = pywinauto.timings.WaitUntilPasses(5,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Required PROPS Input Complete', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()


w_handle = pywinauto.timings.WaitUntilPasses(3,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Economic Analysis', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Close']
ctrl.Click()

w_handle = pywinauto.timings.WaitUntilPasses(3,0.5,lambda:pywinauto.findwindows.find_windows(title=u'PropsPlus.aprbkp - Aspen Properties V7.3 - aspenONE - [Control Panel]')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#Tools->Next
window.TypeKeys('%t')
window.TypeKeys('n')


w_handle = pywinauto.timings.WaitUntilPasses(3,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Properties Plus Complete', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()


print ('Exit Properties')

w_handle = pywinauto.timings.WaitUntilPasses(3,0.5,lambda:pywinauto.findwindows.find_windows(title=u'PropsPlus.aprbkp - Aspen Properties V7.3 - aspenONE - [Control Panel]')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#File->Save
window.TypeKeys('%f')
window.TypeKeys('s')
time.sleep(.5)
# File->Exit
window.edit.TypeKeys("%f")
window.edit.TypeKeys("x")


#This code needed first time and window not disabled TODO 

#w_handle = pywinauto.findwindows.find_windows(title=u'Aspen Properties', class_name='#32770')[0]
#window = pwa_app.window_(handle=w_handle)
#window.SetFocus()
# Don't display anymore
#ctrl = window['&No']
#ctrl.Click()


w_handle = pywinauto.timings.WaitUntilPasses(3,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Physical Properties Configuration', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

time.sleep(1)

print('Now move over chemicals')
#Select the Components List view in All Items
w_handle = pywinauto.timings.WaitUntilPasses(2,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['AfxMDIFrame902']
ctrl.SetFocus()
ctrl = window['TreeView']
ctrl.ClickInput()
ctrl.ClickInput()

window.TypeKeys('{PGUP}')
window.TypeKeys('{DOWN}')
window.TypeKeys('{TAB}')
#Works in both three across or two across width
window.TypeKeys('{RIGHT}')
window.TypeKeys('{RIGHT}')
window.TypeKeys('{DOWN}')
window.TypeKeys('{ENTER}')

           
#Move the four chemicals to the active window
w_handle = pywinauto.timings.WaitUntilPasses(4,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Build Component List - Default', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['List1']
ctrl.SetFocus()
ctrl.Select(0)
ctrl = window['>']
ctrl.Click()
ctrl = window['List1']
ctrl.SetFocus()
ctrl.Select(0)
ctrl = window['>']
ctrl.Click()
ctrl = window['List1']
ctrl.SetFocus()
ctrl.Select(0)
ctrl = window['>']
ctrl.Click()
ctrl = window['List1']
ctrl.SetFocus()
ctrl.Select(0)
ctrl = window['>']
ctrl.Click()
ctrl = window['OK']
ctrl.Click()

#Open solver options and change the non -linear solver from standard to DMO
w_handle = pywinauto.timings.WaitUntilPasses(2,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%r')
window.TypeKeys('v')
w_handle = pywinauto.timings.WaitUntilPasses(5,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Solver Options', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%{DOWN}')
window.TypeKeys('{DOWN}')
window.TypeKeys('{ENTER}')
window.TypeKeys('{ENTER}')


print('Get the needed file')
#Import in the needed file
w_handle = pywinauto.timings.WaitUntilPasses(2,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#File -> Import
window.TypeKeys('%f')
window.TypeKeys('i')


w_handle = pywinauto.timings.WaitUntilPasses(2,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Open', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['ComboBox']
ctrl.SetFocus()
ctrl = window['Edit']
ctrl.SetFocus()

#file name below may change, see note at top

window.edit.TypeKeys(Needed_file)
ctrl = window['&Open']
ctrl.Click()

print('Get the model')
#Get the Adsorber model
w_handle = pywinauto.timings.WaitUntilPasses(6,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Exploring - Simulation']
ctrl.SetFocus()
time.sleep(.5)

#Narrow width for so upcoming ghost mouse script will work
ctrl.MoveWindow(x=0, y=0, width = 200)
ctrl = window['AfxMDIFrame902']
ctrl.SetFocus()

#Assure depth of window so mouse will work
ctrl.MoveWindow(x=0, y=0, height = 238)
ctrl = window['TreeView']
ctrl.SetFocus()
window.TypeKeys('{DOWN}')
window.TypeKeys('{DOWN}')
window.TypeKeys('{DOWN}')
window.TypeKeys('{TAB}')
window.TypeKeys('{RIGHT}')
window.TypeKeys('{ENTER}')
window.TypeKeys('{TAB}')
window.TypeKeys('{TAB}')
window.TypeKeys('{DOWN}')

#Assure depth of window so mouse will work
ctrl = window['AfxMDIFrame903']
ctrl.SetFocus()
ctrl.MoveWindow(x=0, y=238, height = 610)

#Get mouse scripts
ChangeMouse('drag_hollow_fiber.rms')
SetMouseSpeed(10)



#center mouse
w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl.ClickInput()


#Run Ghost Mouse Script to drag Model to Process Flowsheet (Control F2)
ctrl.TypeKeys('^{F2}')

#wait for mouse to do its thing
time.sleep(5) #assume x10 speed

#Rename icon
print ('Rename Icon')
#click on icon to assure menu bar correct
w_handle =pywinauto.timings.WaitUntilPasses(3,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl.TypeKeys('^k') # change icon
ctrl.TypeKeys('^m') # rename icon

w_handle =pywinauto.timings.WaitUntilPasses(5,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Input', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Edit']
ctrl.TypeKeys('M1')
ctrl = window['OK']
ctrl.Click()


# Get Device Variables Table
w_handle =pywinauto.timings.WaitUntilPasses(10,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%l')
window.TypeKeys('m')
window.TypeKeys('{DOWN}')
window.TypeKeys('{ENTER}')



#Set up the fixed device variables
print ('Set up the fixed device variables')
ctrl = window['M1.DeviceVariables Table']
ctrl.SetFocus()
ctrl.TypeKeys('{DOWN}') #skip first row

alphaC02 = 1.0
alphaH2O = .5
alphaN2 = 50
alphaO2 = 50
Ccfct = .51
Dfi = .0004
Dfo = .0006
L = 1
Qed = .12047

window.TypeKeys(str(alphaC02)+ '{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(alphaH2O)+ '{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(alphaN2)+ '{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(alphaO2)+ '{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Ccfct)+ '{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Dfi)+ '{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Dfo)+ '{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(L)+ '{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Qed)+ '{DOWN}')
time.sleep(keywait)
ctrl.Close()

# Move icon to center
print ('Move icon to center')
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%l')
window.TypeKeys('d')
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Find Object', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#Find the icon
ctrl = window['ListBox']
ctrl.Click()
ctrl = window['Find']
ctrl.Click()
window.Close()
#click on icon to assure menu bar correct
w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl.ClickInput()


#change mouse script
ChangeMouse('Stream1.rms')


w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()

#add feed, retentate, permeate streams
window.TypeKeys('^{F2}')

time.sleep(6)#assume x10 speed


#set focus on gas feed specification
print ('set focus on gas feed specification')
w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['feed1.AllVariables Table']
ctrl.SetFocus()


#set Gas feed Values:
print ('Gas feed Values:')
FeedF = 100000
FeedP = 2.0
FeedT =  50
FeedzCO2 =.19
FeedzH2O = .04
FeedzN2 = .72
FeedzO2 = .05

ctrl.TypeKeys('{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys(str(FeedF)+'{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys(str(FeedP)+'{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys(str(FeedT)+'{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys(str(FeedzCO2)+'{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys(str(FeedzH2O)+'{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys(str(FeedzN2)+'{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys(str(FeedzO2)+'{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.Close()



#change mouse script
ChangeMouse('Stream2.rms')


w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()


#add feed, retentate, permeate streams
window.TypeKeys('^{F2}')

time.sleep(2)#assume x10 speed

#set focus on permeate all Variables table
w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['permeate1.AllVariables Table']
ctrl.SetFocus()

PermP= .2
ctrl.TypeKeys('{DOWN 3}')
time.sleep(keywait)
ctrl.TypeKeys(str(PermP)+'{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.Close()

#Run the first simulation
print ('Run the first simulation')
w_handle = pywinauto.timings.WaitUntilPasses(15,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
time.sleep(.5)
window.TypeKeys('{F5}')


#Wait for and then dismiss the done dialog box up to 20 seconds
w_handle = pywinauto.timings.WaitUntilPasses(20,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Error', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()                                         
print ('Dismiss the done dialog box')
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

#reset simulation
print ('Reset simulation')
w_handle = pywinauto.timings.WaitUntilPasses(15,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
time.sleep(.5)
window.TypeKeys('^{F7}')

w_handle = pywinauto.timings.WaitUntilPasses(20,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Reset', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['&Yes']
ctrl.Click()

#run IPsolve script
#center mouse
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl.ClickInput()

w_handle = pywinauto.timings.WaitUntilPasses(15,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%l')
window.TypeKeys('p')
window.TypeKeys('{ENTER}')

w_handle = pywinauto.timings.WaitUntilPasses(20,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Scripting', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
                                             
print ('Dismiss the done dialog box')
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

#change mouse script
ChangeMouse('Stream3.rms')

w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()

#Run Ghost Mouse Script to drag Model to Process Flowsheet (Control F2)
ctrl.TypeKeys('^{F2}')

#wait for mouse to do its thing
time.sleep(6) #assume x10 speed

#Set M2.DeviceVariables Table
w_handle =pywinauto.timings.WaitUntilPasses(10,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)

#Set up the M2 device variables
print ('Set up the M2 device variables')
ctrl = window['M1.DeviceVariables Table']
ctrl.SetFocus()
ctrl.TypeKeys('{DOWN}') #skip first row

alphaC02 = 1.0
alphaH2O = .5
alphaN2 = 50
alphaO2 = 50
Ccfct = .86
Dfi = .0004
Dfo = .0006
L = 1
Qed = .12047

window.TypeKeys(str(alphaC02)+ '{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(alphaH2O)+ '{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(alphaN2)+ '{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(alphaO2)+ '{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Ccfct)+ '{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Dfi)+ '{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Dfo)+ '{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(L)+ '{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Qed)+ '{DOWN}')
time.sleep(keywait)
ctrl.Close()

#change mouse script
ChangeMouse('Stream4.rms')

w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()

#Run Ghost Mouse Script to drag Model to Process Flowsheet (Control F2)
ctrl.TypeKeys('^{F2}')

#wait for mouse to do its thing
time.sleep(6) #assume x10 speed

w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl = window['Sweep.AllVariables Table']
ctrl.SetFocus()



#set Sweep All Variables Values:
print ('Sweep All Variables Values:')
SweepF = 65000
SweepP = 1.3
SweepzCO2 =.0003
SweepzH2O = .01009
SweepzN2 = .78223
SweepzO2 = .20738

ctrl.TypeKeys('{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys(str(SweepF)+'{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys(str(SweepP)+'{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys(str(SweepzCO2)+'{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys(str(SweepzH2O)+'{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys(str(SweepzN2)+'{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys(str(SweepzO2)+'{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.Close()

#change mouse script
ChangeMouse('Stream5.rms')

w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()

#Run Ghost Mouse Script to drag Model to Process Flowsheet (Control F2)
ctrl.TypeKeys('^{F2}')

#wait for mouse to do its thing
time.sleep(2) #assume x10 speed


w_handle = pywinauto.timings.WaitUntilPasses(20,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Scripting', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
                                             
print ('Dismiss the done dialog box')
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

#Check final result
#change mouse script
ChangeMouse('Stream6.rms')

w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()

#Run Ghost Mouse Script to drag Model to Process Flowsheet (Control F2)
ctrl.TypeKeys('^{F2}')

#wait for mouse to do its thing
time.sleep(1) #assume x10 speed

w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Retentate2.AllVariables Table']
ctrl.SetFocus()
ctrl = window['GXWND']
ctrl.SetFocus()
ctrl.TypeKeys('{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys('{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys('{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys('{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys('{DOWN}')
time.sleep(keywait)
ctrl.TypeKeys('{DOWN}')
time.sleep(keywait)


#Make sure cursor highlights selected value of interest
ctrl.TypeKeys('{F2}')
time.sleep(.5)
ctrl = window['Edit']
time.sleep(keywait)

expected = .0176
#Get value highlighted
Properties=ctrl.Texts()
#Convert string to number
actual= float(Properties[1])

print ('Test Results Hollow Fiber Gas Permeation Membrane Simulation '+ Needed_file +' '+ time.asctime())

#Compare
if (actual <= expected):
     print ('Test pass'+' '+'Actual= '+str(actual)+' '+'Expected<= '+str(expected)+ ' Difference= '+str(abs(actual-expected)))
else:
     print ('Test fail'+' '+'Actual= '+str(actual)+' '+'Expected<= '+str(expected)+ ' Difference= '+str(abs(actual-expected)))



