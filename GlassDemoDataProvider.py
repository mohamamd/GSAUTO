import pandas as pd
import os
import xml.etree.ElementTree as ET
from xml.etree.ElementTree  import Element, SubElement

class GlassDemoDataProvider():
    rel_path = "testData/sample.xlsx"
    AxHxvalues=""
    popValues=""
    def getCoordinates(self,key_name,XorY):
        it = ET.iterparse(file("testData/Screenconfig.xml"))
        for ev, el in it:
            if el.tag==key_name:
                keyname= el
                value =int( keyname.get(XorY))
                return value
        return "Not Found"        
    def getFilepaths(self,filename):
        it = ET.iterparse(file("testData/Filepathconfig.xml"))
        for ev, el in it:
            if el.tag==filename:
                keyname= el
                value =keyname.get("path")
                return value
        return "Not Found"         
    def extractvalues(self, df):
        a=str(df).split("Index:")
        val= str(a[-1]).split("[")
        value=str(val[-1]).split("]")
        finalval=str(value[0])
        finalval=finalval.replace(" ","")
        values= finalval.split(",")
        self.popValues = values
        return values
    
    def getvalues(self,xlsheetname):
        script_dir = os.path.dirname(__file__)
        #rel_path = "testData/sample.xlsx"
        abs_file_path = os.path.join(script_dir, self.rel_path)
        df = pd.read_excel(self.rel_path,sheetname=xlsheetname, index_col=0,parse_cols = 0,)
        self.AxHxvalues = self.extractvalues(df)

        return self.AxHxvalues


        
    def __init__(self,rel_path="testData/sample.xlsx"):
        self.rel_path=rel_path
        
##scroll_down=(str(GlassDemoDataProvider().getCoordinates('scroll_down',"x")),str(GlassDemoDataProvider().getCoordinates('scroll_down',"y")))
##standby=(str(GlassDemoDataProvider().getCoordinates('standby',"x")),str(GlassDemoDataProvider().getCoordinates('standby',"y")))
##open_fullscreen=(str(GlassDemoDataProvider().getCoordinates('open_fullscreen',"x")),str(GlassDemoDataProvider().getCoordinates('open_fullscreen',"y")))
##exit_fullscreen=(str(GlassDemoDataProvider().getCoordinates('exit_fullscreen',"x")),str(GlassDemoDataProvider().getCoordinates('exit_fullscreen',"y")))
##Guide=(str(GlassDemoDataProvider().getCoordinates('Guide',"x")),str(GlassDemoDataProvider().getCoordinates('Guide',"y")))
##exitkey=(str(GlassDemoDataProvider().getCoordinates('exitkey',"x")),str(GlassDemoDataProvider().getCoordinates('exitkey',"y")))
##down=(str(GlassDemoDataProvider().getCoordinates('down',"x")),str(GlassDemoDataProvider().getCoordinates('down',"y")))
##menu=(str(GlassDemoDataProvider().getCoordinates('menu',"x")),str(GlassDemoDataProvider().getCoordinates('menu',"y")))
##up=(str(GlassDemoDataProvider().getCoordinates('up',"x")),str(GlassDemoDataProvider().getCoordinates('up',"y")))
##select=(str(GlassDemoDataProvider().getCoordinates('select',"x")),str(GlassDemoDataProvider().getCoordinates('select',"y")))
##right=(str(GlassDemoDataProvider().getCoordinates('right',"x")),str(GlassDemoDataProvider().getCoordinates('right',"y")))
##left=(str(GlassDemoDataProvider().getCoordinates('left',"x")),str(GlassDemoDataProvider().getCoordinates('left',"y")))
##scroll_up=(str(GlassDemoDataProvider().getCoordinates('scroll_up',"x")),str(GlassDemoDataProvider().getCoordinates('scroll_up',"y")))
##key_one=(str(GlassDemoDataProvider().getCoordinates('key_one',"x")),str(GlassDemoDataProvider().getCoordinates('key_one',"y")))

path=GlassDemoDataProvider().getFilepaths('Report_src')
print path
print "done"
##GlassDemoDataProvider().editTagInReportconfig()
##print getCoordinates('menu_key',"y")

##A6H1 = GlassDemoDataProvider("testData/sample.xlsx").getvalues("A6H1")
##A7H1 = GlassDemoDataProvider("testData/sample.xlsx").getvalues("A7H1")
##A7H2 = GlassDemoDataProvider("testData/sample.xlsx").getvalues("A7H2")
##
##print A6H1
##print A7H1
##print A7H2
#print GlassDemoDataProvider().getCoordinates('key_two',"y")

        
