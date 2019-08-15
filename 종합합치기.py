import uno
# get the uno component context from the PyUNO runtime
localContext = uno.getComponentContext()

# create the UnoUrlResolver
resolver = localContext.ServiceManager.createInstanceWithContext(                            "com.sun.star.bridge.UnoUrlResolver", localContext)

# connect to the running office
ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
smgr = ctx.ServiceManager

# get the central desktop object
DESKTOP =smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
###################################
#import root folder path
rfolder = 'C:/pycon/T1.integrate/ex2.pivot/'
# rfolder = 'C:/Users/kyoohoA/Desktop/pycon/T1.integrate/ex.pivot/'

#calling to calc model
turl = 'file:///'+rfolder+'종합.ods'
tmodel = DESKTOP.getCurrentComponent()
tmodel = DESKTOP.loadComponentFromURL(turl,"_blank",0,() )
if not hasattr(tmodel, "Sheets"):
 tmodel = DESKTOP.loadComponentFromURL("private:factory/scalc","_blank", 0, () )
tsheet = tmodel.Sheets.getByIndex(0)
tendrow = 1

#loop Start
filenames = ['보험팀.ods','사업지원팀.ods','생산지원팀.ods','세무팀.ods','인사팀.ods','총무팀.ods','회계팀.ods']
for filename in filenames:
    #calling original model
    urladress = 'file:///'+rfolder+filename
    fmodel = DESKTOP.loadComponentFromURL(urladress,"_blank",0,() )   
    fsheet = fmodel.Sheets.getByIndex(0)
    cursor = fsheet.createCursor()
    #goto the last used cell
    cursor.gotoEndOfUsedArea(True)
    #grab that positions "coordinates"
    faddress = cursor.RangeAddress
    fendrow = faddress.EndRow
    fcellname = 'A2:E'+str(fendrow)
    fRange = fsheet.getCellRangeByName(fcellname)
    fArray = fRange.getDataArray()
    #paste to file
    tendrow += 1
    tcellname = 'A'+str(tendrow)+':E'+str(tendrow+fendrow-2)
    tendrow += fendrow -2
    tRange = tsheet.getCellRangeByName(tcellname)
    tRange.setDataArray(fArray)

#################################
#end of loop
from com.sun.star.beans import PropertyValue
args = (PropertyValue('FilterName', 0, 'MS Excel 97', 0),)
tmodel.storeAsURL(turl, args)
