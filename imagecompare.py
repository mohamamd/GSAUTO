from PIL import Image
from PIL import ImageChops
import math
import operator
import os
import shutil
from pytesseract import image_to_string
import pytesseract
import os.path
import openpyxl
class imagecompare():
    spectrunverificationlist=['TV Shows','Included with','All Subscriptions','HBO','Primetime','FOX','THOUSANDS OF SHOWS ON DEMAND','OF SHOWS ON DEMAND','My Subscriptions','Primetime Free']
    im1=""
    im2=""
    continueTest = True
    currentSlotID= ''
    currentHeadendhub=''
    result='PASS'
    failedfetures='none'
    currentReportRow=1
    previousSlotID="0"
    

    PATH='testData/Glas Automation Status.xlsx'
    
    pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract.exe'
    def resetReportValues(self):
        self.result="PASS"
        self.failedfetures=""
        self.currentHeadendhub=""
        self.currentSlotID=""

    def setCurrentHEHubandSlotID(self ,currentHeadendhub,SlotID):
        """ Setting the current slot id is executing, value will be recieved runtime from reobot. slot id should be passed.Current slot will be updated in in script"""
        self.currentSlotID =SlotID
        self.currentHeadendhub = currentHeadendhub
    def getCurrentSlotID(self):
        """ returning the current slot id is executing to robot, no input required"""
        return self.currentSlotID
    
    def getvalue(self):
        """ calculating RMS value diffrence of pixels in 2 screenshots, returning rms value. no input should be passed """
        h1 = Image.open(self.im1).histogram()
        h2 = Image.open(self.im2).histogram()
        rms = math.sqrt(reduce(operator.add,map(lambda a,b: (a-b)**2, h1, h2))/len(h1))
        print 'rms value is' , rms
        return rms
    def match70percent(self,im1,im2):
        """ this method can be used for future. this functioncan can match the screen shots by a given value, both screen shots objects should be passed. Returning True/False based on rms of dffreence of pixel values"""
        self.im1=im1
        self.im2=im2
        x=(100-self.getvalue())
        
        if ( x> 70) :
            return True
        else: return False
    def IsAvAvailable(self,im1,im2):
        """This API is used to verify the AV from 2 given screen shots by comaparing it.  both screen shots objects should be passed. Returning True/False based on rms of dffreence of pixel values"""
        self.im1=im1
        self.im2=im2
        x=(self.getvalue())
        print 'value of x in isAvailable', x
        if ( x < 150) :
            #updating av verification falure on report variebls #
            self.result="FAIL"
            if (("GLAS(ERROR IN RETRIEVING DATA)" in self.failedfetures) or("BOX REBOOTING" in self.failedfetures )or ("GLAS(NOIR/NOSIGNAL/UNSUPPORTEDSIGNAL)" in self.failedfetures ) ):
                pass
            else:
                self.failedfetures=self.failedfetures +" NO VIDEO"
        
            
            #updated
            return False
        else: return True
        
    def getlastscreenshot(self,prefix,screenshotfolder="screenshot"):
        """ This API is used to get last scren shot captured. Screen shot prefix name,screenshot saved path should input. Last screen shot name will be returned"""
        filepattern=screenshotfolder+"/"+prefix
        listOfFiles=os.listdir(screenshotfolder)
        last=0
        
        count=0
        number=0
        #print prefix, "is prefix"
        if len(listOfFiles)<1:
            print "no screen shot to compare"
            return False
        for  x in listOfFiles:
            #print x, prefix
            if prefix in (x):
                #print x, prefix," matching with prefix"
                val= x.split(prefix)
                val= str(val[1]).split(".")
                number= int(val[0])
                #print number
                if number > last :
                    last=number
                    count=count+1

        #print count
        if count < 1 :
            return False
        lastfile=filepattern+str(last)+".png"
        print "last screen shot is ", lastfile
        return lastfile
        
    def last2ScreenshotsCompareForAV(self,prefix,screenshotfolder="screenshot"):
        """This function is used to get last 2 screen shot captured, this can be used only inside file.This function should not expose to outer world. Screen shot prefix name,screenshot saved path should input """
        filepattern=screenshotfolder+"//"+prefix
        listOfFiles=os.listdir(screenshotfolder)
        last=0
        secondlast=0
        count=0
        number=0
        #print prefix, "is prefix"
        if len(listOfFiles)<2:
            #print "no screen shot to compare"
            return False
        for  x in listOfFiles:
            #print x, prefix
            if prefix in (x):
                #print x, prefix," matching with prefix"
                val= x.split(prefix)
                val= str(val[1]).split(".")
                number= int(val[0])
                #print number
                if number > last :
                    secondlast=last
                    last=number
                    count=count+1
                elif number > secondlast :
                    secondlast=number
                    
        #print count
        if count < 2 :
            return False
        file1=filepattern+str(last)+".png"
        file2=filepattern+str(secondlast)+".png"
        print "comparing ",file1 ," and  ",file2 
        return self.IsAvAvailable(file1,file2)
    
    def DeleteAllOldScreenshotsin(self,screenshotfolder="screenshot"):
        """ deletes all the screen shot caputured by robot.Currently not used in Script file, screenshots saved path should input"""
        for the_file in os.listdir(screenshotfolder):
            file_path = os.path.join(screenshotfolder, the_file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                #elif os.path.isdir(file_path): shutil.rmtree(file_path)
            except Exception as e:
                print(e)
        print "screen shot folder cleared"
    def verifyBoxAV(self,prefix,screenshotfolder="screenshot"):
        """ This API used to check any strean issue is encountered while execution. Screen shot prefix name,screenshots saved path should input"""
        lastscreenshot=self.getlastscreenshot(prefix,screenshotfolder)
        x= image_to_string(Image.open(lastscreenshot), lang='eng')
        print "contents of last screen shot is " ,x
        if ('Preparing Stream' in x) or ('Checking Stream' in x) or ('... on' in x)or ('Checking Power Socket' in x) or ('Checking STB' in x) or ('Error' in x):
            #updating in report vriebles for reporting
            self.result="FAIL"
            self.failedfetures="GLAS(ERROR IN RETRIEVING DATA)"
            #update done
            return False 
        else:
            return True
    def onDemandVerification(self,prefix,screenshotfolder="screenshot"):
        """Verifying ON DEMAND screen. Screen shot prefix name,screenshots saved path  should input.returns True/Flase"""
        lastscreenshot=self.getlastscreenshot(prefix,screenshotfolder)
        #print "last screenshot is" ,lastscreenshot
        x= image_to_string(Image.open(lastscreenshot), lang='eng')
        print "contents of last screen shot is " ,x
        flag1 = False
        count = x.count('Free')
        if  count>1:
            flag1 = True
            print ('found free multiple times')
        
        elif ('TV Shows' in x) or ('Included with' in x) or (('FOX' in x) and ('FOX 2' not in x ))or ('All Movies' in x) or (('More' in x)and('Choices' in x)):
            flag1 = True
            print "found TV Shows or Included with or FOX"
        elif 'Welcome to On Demand' in x:
            
            flag1 = True
            print "found Welcome to On demand"        
        elif ('All Subscriptions' in x) or ('All Networks'in x):
            flag1 = True
            print "found All Subscriptions or All Networks"
        elif ('HBO' in x) or ('Free previews' in x)or ('FX' in x):
            flag1 = True
            print "found HBO or Free previews or FX"
        elif ('Adult' in x) or  ('Free Movies' in x) or ('Canoe' in x):
            flag1 = True
            print "found Adult or Free or Free Movies"
        print flag1 ,"after check"
        if (not flag1):
            self.result="FAIL"
            if (("GLAS(ERROR IN RETRIEVING DATA)" in self.failedfetures) or("BOX REBOOTING" in self.failedfetures )or ("GLAS(NOIR/NOSIGNAL/UNSUPPORTEDSIGNAL)" in self.failedfetures ) ):
                pass
            else:
                self.failedfetures=self.failedfetures +" VOD FAILED"  
        return flag1
    def GuideVerification(self,prefix,slotID="0",screenshotfolder="screenshot" ):
        """ API to Verify Guide screen , writing to the file if fails and returning status to robot.Screen shot prefix name,screenshots saved path  and current slot id[optional] should input"""
        #slotID='under devolopment'
        lastscreenshot=self.getlastscreenshot(prefix,screenshotfolder)
        x= image_to_string(Image.open(lastscreenshot), lang='eng')
        print "contents of last screen shot is " ,x
        flag = False
        count = x.count('On Demand')
        if  count>0:
            flag = True
            print "found on demand multiple times"
        elif ('002' in x) or ('003' in x) or ('004' in x) or ('News' in x) :
            flag =True
            print "found 002 or 003 or 004 or news"
        elif ('Fox 2 News' in x) or ('TCT Partnership' in x) or ('Paid Program' in x) or ('Partnership' in x) or (('Announced' in x)and ('To Be' in x)):
            flag = True
            print "found Fox 2 News or TCT Partnership or Paid Program or Partnership"
        if (not flag):
            temp=screenshotfolder
            if '/' in temp:
                temp=temp.split("/")
            elif '\\' in temp: 
                temp=temp.split("\\")
            current_execution_folder=temp[0]
            guideVerificationFailed=open(current_execution_folder+'/guideVerificationFailed.txt','a')
            guideVerificationFailed.write(slotID)
            guideVerificationFailed.write ("\n")
            guideVerificationFailed.close()
            self.result="FAIL"
            if (("GLAS(ERROR IN RETRIEVING DATA)" in self.failedfetures) or("BOX REBOOTING" in self.failedfetures )or ("GLAS(NOIR/NOSIGNAL/UNSUPPORTEDSIGNAL)" in self.failedfetures ) ):
                pass
            else:
                self.failedfetures=self.failedfetures +" GUIDE FAILED"            

        return flag
    
    def ReservationVerification(self,prefix,slotID="0", screenshotfolder="screenshot"):
        """Verifying reservation status,writing to the file if reserved and returnin status to robot . Screen shot prefix name,screenshots saved path  and current slot id[optional] should input"""
        temp=screenshotfolder
        if '/' in temp:
            temp=temp.split("/")
        elif '\\' in temp: 
            temp=temp.split("\\")
        current_execution_folder=temp[0]
        lastscreenshot=self.getlastscreenshot(prefix,screenshotfolder)
        x= image_to_string(Image.open(lastscreenshot), lang='eng')
        print "contents of last screen shot is " ,x
        flag = True
        print 'checking reservation staus of ', slotID , ' from the screenshot'
        if (('Checking' in x)and ('Reservation' in x)) and (('Available' not in x) and ('Quick Reserve' not in x)):
            flag = False
            print slotID , "is reserved"
            ReservedBoxes=open(current_execution_folder+'/ReservedBoxes.txt','a')
            
            ReservedBoxes.write(slotID)
            ReservedBoxes.write ("\n")
            ReservedBoxes.close()
            #updating report variebles for reporting
            self.result="FAIL"
            self.failedfetures="RESERVED BOX"
            #"RESERVED BOX" will be reported.
        elif(('Reserved By' not in x) and ('Available' not in x) and ('Quick Reserve' not in x)):#monitor. dependency may there in glas2.2
            flag= None
        return flag
    def findWatchOption(self,prefix,screenshotfolder="screenshot"):
            """This API verifying watch option in screen.Screen shot prefix name,screenshots saved path  should input"""
            lastscreenshot=self.getlastscreenshot(prefix,screenshotfolder)
            x= image_to_string(Image.open(lastscreenshot), lang='eng')
            print "contents of last screen shot is " ,x
            if ('WATCH' in x ) or ('watch' in x ) or ('Watch' in x):
                return True
            else: false
            
    def verifyNOSIGNAL_UNSUPPORTEDSIGNAL(self,prefix,screenshotfolder="screenshot"):
            """ checking for any No Signal ,Unsupported signal and booting issues. Screen shot prefix name,screenshots saved path  should input"""
            lastscreenshot=self.getlastscreenshot(prefix,screenshotfolder)
            x= image_to_string(Image.open(lastscreenshot), lang='eng')
            print "contents of last screen shot is " ,x
            if ('No Signal' in x) or ('Unsupported Signal' in x):
                return False
            elif ('Unable to Start' in x) or ('Digital Receiver Starting' in x):
                self.result="FAIL"
                self.failedfetures= "BOX REBOOTING"
                return 'Booting'
            
            else: return True
    def reportNOSIGNAL_UNSUPPORTEDSIGNAL_NOIR(self):
            self.result="FAIL"
            self.failedfetures="GLAS(NOIR/NOSIGNAL/UNSUPPORTEDSIGNAL)"
                
            
    def returnTrue(self,anything='anything'):
        """ mechanism to assign True in robot file, returns true alwas"""
        return True
    def getcontinueTest(self):
        """ This API will confirm weather execution should be stopped or not. Returns True or False"""
        return self.continueTest
    def setcontinueTest(self,Bool):
        """ This API setting value of continueTest which decides the execution flow. input True / False """
        self.continueTest=Bool
    def getBoxType(self,prefix,screenshotfolder="screenshot"):
            """ This API is used to confirm the box type.  Screen shot prefix name,screenshots saved path  should input. box type will be returned """
            lastscreenshot=self.getlastscreenshot(prefix,screenshotfolder)
            x= image_to_string(Image.open(lastscreenshot), lang='eng')
            print "contents of last screen shot is " ,x
            boxtype='none'
            for element in self.spectrunverificationlist :
                if (element in x):
                    print 'found ',element
                    boxtype ='spectrum'
                    break
            if ('DON\'T SHOW AGAIN' in x )and( boxtype =='spectrum'):
                boxtype ='First Time Spectrum' 

            return boxtype
    def detectnormalScreen(self,prefix,screenshotfolder="screenshot"):
            """ checking whether execution happening in full screen or in normal screen Screen shot prefix name,screenshots saved path  should input. returns True/Flase"""
            lastscreenshot=self.getlastscreenshot(prefix,screenshotfolder)
            x= image_to_string(Image.open(lastscreenshot), lang='eng')
            print "contents of last screen shot is " ,x
            flag = False
            if 'Lab Glas' in x:
                flag= True
            elif ('Available' in x) or ('Preparing' in x) or ('Quick Reserve' in x):
                flag = True
            elif 'Welcome to Glas' in x:
                flag = True
            elif('Stream' in x )or ('Checking' in x)or('Select Device' in x) or ('Reservation' in x) or ('Successfully' in x)or('Ended' in x) or ('Early' in x):
                flag= True
            elif ('Reserved' in x) or ('Until' in x) or ('IST' in x):
                flag = True
            return flag

    def VerifySearchScreen(self,prefix,screenshotfolder="screenshot"):
            """ verifying execution hapening in search screen.  Screen shot prefix name,screenshots saved path  should input. Returns True or Flase"""
            lastscreenshot=self.getlastscreenshot(prefix,screenshotfolder)
            x= image_to_string(Image.open(lastscreenshot), lang='eng')
            print "contents of last screen shot is " ,x
            flag = False
            if ('Search' in x )or ('Search for TV Shows' in x ) or ('Movies' in x )or ('and Networks' in x )or (('TV Shows' in x )and ('Movies' in x )and ('Networks' in x)):
                print '''('Search' in x )or ('Search for TV Shows' in x ) or ('Movies' in x )or ('and Networks' in x )or (('TV Shows' in x )and ('Movies' in x )and ('Networks' in x)):'''
                flag= True
            elif('Recent Searches' in x)or ('Search by title' in x)or ('keyword or' in x) or ('actor name'in x):
                flag = True
                print '''('Recent Searches' in x)or ('Search by title' in x)or ('keyword or') or ('actor name'in x):'''
            elif ('Delete' in x) or ('Clear' in x) :
                flag= True
            return flag
    def verifyEmpireSynopsysScreen(self,prefix,screenshotfolder="screenshot"):
            """Verifying execution happening in synopsis screen of EMPIRE, Screen shot prefix name,screenshots saved path  should input.returns True/Flase"""
            lastscreenshot=self.getlastscreenshot(prefix,screenshotfolder)
            x= image_to_string(Image.open(lastscreenshot), lang='eng')
            print "contents of last screen shot is " ,x
            flag = False
            if ('Empire' in x) or ('season' in x) or ('Available on'in x) or ('E4 Cupid Kills'in x ) or('Sound & Fury' in x) or ('Watch' in x):
                flag= True
                print '''('Empire' in x) or ('season' in x) or (' Available on'in x) or ('E4 Cupid Kills') or ('Watch' in x)'''
            elif ('Cupid Kills' in  x) or ('Lucious starts seeing' in x) or ('Angelo as a threat' in x) or ('Jamal'in x) or ('takes his first' in x):
                flag= True
                print '''('Cupid Kills' in  x) or ('Lucious starts seeing' in x) or ('Angelo as a threat' in x) or ('Jamal'in x) or ('takes his first' in x):'''
            elif  ('big step towards recovery' in x) or ('and faces Freda' in x ) or ('return causes' in x) or ('Hakeem' in x) or ('to have' in x):
                flag = True
                print '''('big step towards recovery' in x) or ('and faces Freda' in x ) or ('return causes' in x) or ('Hakeem starts' in x) or ('to have' in x):'''
            elif ('Jamal' in x) or('Lucious' in x) or ('Tiana' in x):
                flag= True
            return flag
    def searchForWatch(self,prefix,screenshotfolder="screenshot"):
            """ Verifying weather asset can be watched by some direct clicks. Screen shot prefix name,screenshots saved path  should input.returns True/Flase"""
            lastscreenshot=self.getlastscreenshot(prefix,screenshotfolder)
            x= image_to_string(Image.open(lastscreenshot), lang='eng')
            print "contents of last screen shot is " ,x
            flag = False
            if ('Watch ' in x)or ('watch ' in x) or ('WATCH ' in x ): #space is needed here otherwise will be true as watchlist is available in screen
                flag= True
            elif('restart' in x)or ('resume' in x)or ('Resume' in x) or ('series info' in x ):
                flag = True
            return flag
    def reportVODPlaybackFailure(self):
            self.result="FAIL"
            self.failedfetures=self.failedfetures + " VOD PLAYBACK FAILED"
                
    def verifyBrowseEpisode(self,prefix,screenshotfolder="screenshot"):
            """ checking fothe string Browse Episoeds i n screen. Screen shot prefix name,screenshots saved path  should input.returns True/Flase"""
            lastscreenshot=self.getlastscreenshot(prefix,screenshotfolder)
            x= image_to_string(Image.open(lastscreenshot), lang='eng')
            print "contents of last screen shot is " ,x
            flag = False
            if ('BROWSE' in x)or ('EPISODES' in x):                 
                return True
            else:
                return False
    def verifyAvailableToWatchNow(self,prefix,screenshotfolder="screenshot"):
            """checking for the string available to watch now. Screen shot prefix name,screenshots saved path  should input.returns True/Flase"""
            lastscreenshot=self.getlastscreenshot(prefix,screenshotfolder)
            x= image_to_string(Image.open(lastscreenshot), lang='eng')
            print "contents of last screen shot is " ,x
            flag = False
            if ('Available to' in x)or ('Watch Now' in x):                 
                return True
            else:
                return False
    def ReportTestStatus(self, executionMode):
        if (executionMode !='Debugg'):
            imagecompare.currentReportRow = imagecompare.currentReportRow + 1
            #ReportGenarator().editMasterDatainExel("C","3","PASS")
            self.editMasterDatainExel("A",imagecompare.currentReportRow,self.currentHeadendhub)
            self.editMasterDatainExel("B",imagecompare.currentReportRow,self.currentSlotID)
            self.editMasterDatainExel("C",imagecompare.currentReportRow,self.result)        
            self.editMasterDatainExel("D",imagecompare.currentReportRow,self.failedfetures)
            imagecompare.previousSlotID=self.currentSlotID
            if (self.result =='PASS'):
                self.updateReportOverview('PASS')
            elif(self.failedfetures=="RESERVED BOX" or self.failedfetures=="BOX REBOOTING"):
                self.updateReportOverview('SKIPPED')
            elif(self.failedfetures=="GLAS(NOIR/NOSIGNAL/UNSUPPORTEDSIGNAL)" or self.failedfetures=="GLAS CRASHED" or self.failedfetures=="GLAS(ERROR IN RETRIEVING DATA)"):
                self.updateReportOverview('GLAS_ISSUE')
            else:
                self.updateReportOverview('STB_ISSUE')
                
            
            print (self.currentHeadendhub,self.currentSlotID, self.result,self.failedfetures)
            print "reporting for "+self.currentSlotID+ "is done"
            
            print imagecompare.currentReportRow
        else: pass
    
    def checkFile(self,PATH):
        if os.path.isfile(PATH) and os.access(PATH, os.R_OK):
            return True
        else:
            return False 
    def readDatafromExcel(self,sheet_name,row,column):
        if self.checkFile :
            xfile = openpyxl.load_workbook(self.PATH)
            sheet = xfile[sheet_name]
            return (sheet[str(row)+str(int(column))].value)

    def checkPreviousRecord(self,SlotID,executionMode):
        if (executionMode !='Debugg'):
            if(SlotID=="0" or SlotID==imagecompare.previousSlotID):
                pass
            else:
                self.failedfetures="GLAS CRASHED"
                self.result="FAIL"
                self.ReportTestStatus(executionMode)
        else: pass            

    def editMasterDatainExel(self,column,row,data):
        """Send data in the format ROW,column,data ex: ("C","3","PASS") """
        if self.checkFile :
            xfile = openpyxl.load_workbook(self.PATH)
            sheet = xfile.get_sheet_by_name('Master Data')
            sheet[str(column)+str(int(row))]=data.strip()
            xfile.save(self.PATH)
            
            return True
        else:
            print "testData/Glas Automation Status.xlsx file is missing"
            return False
    def updateReportOverview(self,content="\n"):
        reportOverView=open("testData/reportOverView.txt",'a+')
        reportOverView.write(content)
        reportOverView.write("\n")
        reportOverView.close()
    def __init__(self):
                pass
    def debug(self):
        pass

#imagecompare().debug()
