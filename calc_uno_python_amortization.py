#!/usr/bin/env python3
#
# calc_uno_python_amortization.py
#
# Python file to launch a Calc spreadsheet and build an Amortization sheet.
# Save and run this python script in "My Macros":
# ~/.config/libreoffice/4/user/Scripts/python/
#
# OR, potentially run in, "LibreOffice Macros", but priv issues may exists etc.:
# /usr/lib/libreoffice/share/Scripts/python/
#
# May be saved and run from other folders, so long as a copy also exists in
# ~/.config/libreoffice/4/user/Scripts/python/ for the callbacks.
#
# The two callbacks display in Control Properties --> Events, as:
# calc_uno_python_amortization.py$cb_scrollbar_mouse_up (user, Python)
# calc_uno_python_amortization.py$cb_scrollbar_adjust (user, Python)
#
# Ian Stewart
# 2021-09-02 CC0
# Tested with Python 3.8.10 and LibreOffice Version: 6.4.7.2 on 
# Ubuntu Mate 20.04.
#
import uno
import sys
import os

# Constants

#print(os.getcwd()) # /home/ian/.config/libreoffice/4/user/Scripts/python
#print(sys.argv[0]) # calc_uno_python_amortization.py
#print(os.sep) # / (linux)

# Place the created Spreadsheet is in same directory as the python script.
# This should aid call-backs to slider controls.
# Get name of python file, replace extension as ods.
FILE_NAME = os.path.splitext(sys.argv[0])[0] + '.ods'
# Python 3.9 will allow path.removesuffix('.py') + '.ods'
#print(FILE_NAME)  # calc_uno_python_amortization.ods

# For when run from another doc.
if FILE_NAME == '.ods':
    FILE_NAME = 'calc_uno_python_amortization.ods'

FILE_PATH_NAME = os.getcwd() + os.sep + FILE_NAME
#print(FILE_PATH_NAME)  # /home/ian/.config/libreoffice/4/user/Scripts/python/calc_uno_python_amortization.ods

FILE_URL = 'file://' +  FILE_PATH_NAME
#print(FILE_URL)  # file:///home/ian/.config/libreoffice/4/user/Scripts/python/calc_uno_python_amortization.ods


# OFFSET is used for Row on Spreadsheet where Table commences
OFFSET = 10

from com.sun.star.beans import PropertyValue
# Get UNO structures.
from com.sun.star.awt import Size
from com.sun.star.awt import Point

from com.sun.star.drawing.FillStyle import SOLID
from com.sun.star.drawing.LineJoint import MITER
from com.sun.star.drawing.LineStyle import DASH
from com.sun.star.drawing.LineStyle import SOLID as LINE_SOLID
from com.sun.star.chart.ChartLegendPosition import BOTTOM
from com.sun.star.table.CellHoriJustify import CENTER
from com.sun.star.awt.FontWeight import NORMAL
from com.sun.star.awt.FontWeight import BOLD   
    

def main_initialize_not_embedded():
    """ Create a Calc document and return desktop and doc """
    # get the uno component context from the PyUNO runtime
    localContext = uno.getComponentContext()
    # create the UnoUrlResolver
    resolver = localContext.ServiceManager.createInstanceWithContext(
	    "com.sun.star.bridge.UnoUrlResolver", localContext )
    
    try:    
        # connect to the running office
        ctx = resolver.resolve( 
            "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
    except Exception as e:
        if e.typeName == "com.sun.star.connection.NoConnectException":
            print("Error: No Connection to LibreOffice.")
            # .NoConnectException: Connector : couldn't connect to socket (Connection refused)
            print('The following command must have been executed from a separate terminal window:')
            #print('$ soffice "-accept=socket,host=localhost,port=2002;urp;"') <-- deprecated
            #print('$ soffice "--accept=socket,host=localhost,port=2002;urp;"')
            print('$ soffice "--accept=socket,host=localhost,port=2083;urp;StarOffice.ServiceManager"')
            print()            
            
            sys.exit("Exiting...")
        else:
            print("Error type:", e.typeName)
            sys.exit("Exiting...") 
    
    # https://wiki.openoffice.org/wiki/Documentation/DevGuide/ProUNO/Characteristics_of_the_Interprocess_Bridge
    # "uno:socket,host=localhost,port=2002;urp,Negotiate=0,ForceSynchronous=0;StarOffice.ServiceManager"
    smgr = ctx.ServiceManager
    # get the central desktop object
    desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
    
    # access the current document. i.e. the model
    #doc = desktop.getCurrentComponent()

    # Create a Calc document
    doc = desktop.loadComponentFromURL( "private:factory/scalc","_blank", 0, () )

    return desktop, doc
    
    
def setup_cell_format(doc, sheet):
    """ Setup the format of some cells. """
    local_settings = uno.createUnoStruct( "com.sun.star.lang.Locale" ) 

    local_settings.Language = "en"
    #local_settings.Country = "us"
    local_settings.Country = "nz"
     
    number_formats = doc.NumberFormats
    
    # Dollars with no cents
    number_format_string = "$#,##0"
    number_format_id = number_formats.queryKey(number_format_string, local_settings, True)
    if number_format_id == -1:
        number_format_id = number_formats.addNew(number_format_string, local_settings)

    # Principal Amount.
    cell = sheet.getCellByPosition(1,1)
    cell.NumberFormat = number_format_id
    # Total Paid
    cell = sheet.getCellByPosition(1,6)
    cell.NumberFormat = number_format_id        
    # Balance column
    cell_range = sheet.getCellRangeByPosition (3, OFFSET, 3, OFFSET + 480)
    cell_range.NumberFormat = number_format_id
    
    # Dollars with cents
    number_format_string = "$#,##0.00"
    number_format_id = number_formats.queryKey(number_format_string, local_settings, True)
    if number_format_id == -1:
        number_format_id = number_formats.addNew(number_format_string, local_settings)

    # Monthly Amount.
    cell = sheet.getCellByPosition(1,7)
    cell.NumberFormat = number_format_id

    # Currency for columns
    cell_range = sheet.getCellRangeByPosition (1, OFFSET, 2, OFFSET + 480)
    cell_range.NumberFormat = number_format_id    
    
    # Horizontally justify column headings
    #oRange.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER changes to:
    cell_range = sheet.getCellRangeByPosition (0, OFFSET-1, 3, OFFSET-1)    
    cell_range.HoriJustify = CENTER
        
    # Make the Heading Bold
    # cell.CharWeight = com.sun.star.awt.FontWeight.BOLD changes to: 
    cell = sheet.getCellByPosition(0,0)
    cell.CharWeight = BOLD    


def setup_slider_initial(doc, sheet):
    """ Insert the Form control sliders """
    draw_page = sheet.DrawPage
    
    # A form is required to hold the Scrollbars.
    form = doc.createInstance("com.sun.star.form.component.Form")
    form.Name = "Form1"
    forms = draw_page.Forms
    forms.insertByIndex(0, form) # Works OK with python
    
    # Insert headings in column A
    heading_list = ["Amortization", "Principal", "Interest P.A.", "Term Years", 
            "Total Months", "", "Total Paid:", "Month Pay:"]
    for i in range(0, len(heading_list)):
        cell = sheet.getCellByPosition(0, i)
        cell.String = heading_list[i]

    # Headings in row 10, index 9.
    heading_list = ["Month", "Interest", "Principal", "Balance"]
    for i in range(0, len(heading_list)):
        cell = sheet.getCellByPosition(i, 9)
        cell.String = heading_list[i]
       
    # Create 3 x Scrollbars. Give each a unique name and set background colour.
    # Use ControlShape to position and size each one. Use global Point and Size
    for i in range(0, 3):
        scrollbar = doc.createInstance("com.sun.star.form.component.ScrollBar")
        scrollbar.BackgroundColor = 0xC8C8FF
        scrollbar.Name = "ScrollBar_" + str(i)

        forms[0].insertByName(scrollbar.Name, scrollbar) 
        
        shape = doc.createInstance ( "com.sun.star.drawing.ControlShape" )
        shape.Control = scrollbar
        draw_page.add(shape)
    
        shape.setPosition(Point(7000, i * 450 + 450 ))
        shape.setSize(Size(6000, 450))

    # Loan Amount. Set Scrollbar_0 parameters. Use Name instead of Index.
    # Value is multiplied by 10000
    if forms[0].hasByName("ScrollBar_0"):        
        scrollbar_temp = forms[0].getByName("ScrollBar_0")        
        #scrollbar_temp.Name
        scrollbar_temp.ScrollValueMin = 1
        scrollbar_temp.ScrollValueMax = 100
        scrollbar_temp.DefaultScrollValue = 3
        scrollbar_temp.LineIncrement = 1    # Small change
        scrollbar_temp.BlockIncrement = 10 # Large Change
        #scrollbar_temp.Orientation
        #scrollbar_temp.Tag
        #scrollbar_temp.addEventListener

    # Interest P.A. Set Scrollbar_1 parameters. Use Name instead of Index.
    # Value is multiplied by 0.1
    if forms[0].hasByName("ScrollBar_1"):
        scrollbar_temp = forms[0].getByName("ScrollBar_1")
        #scrollbar_temp.Name
        scrollbar_temp.ScrollValueMin = 1
        scrollbar_temp.ScrollValueMax = 100
        scrollbar_temp.DefaultScrollValue = 30
        scrollbar_temp.LineIncrement = 1    # Small change
        scrollbar_temp.BlockIncrement = 10 # Large Change
        #scrollbar_temp.Orientation
        #scrollbar_temp.Tag
        #scrollbar_temp.addEventListener
    
    # Term in Years. Set Scrollbar_2 parameters. Use Name instead of Index.
    # Value is multiplied by 1
    if forms[0].hasByName("ScrollBar_2"):
        scrollbar_temp = forms[0].getByName("ScrollBar_2")
        #scrollbar_temp.Name        
        scrollbar_temp.ScrollValueMin = 1
        scrollbar_temp.ScrollValueMax = 40
        scrollbar_temp.DefaultScrollValue = 4
        scrollbar_temp.LineIncrement = 1    # Small change
        scrollbar_temp.BlockIncrement = 10 # Large Change
        #scrollbar_temp.Orientation
        #scrollbar_temp.Tag
        #scrollbar_temp.addEventListener   
    
    # Setup the adjust and mouse-up listeners for the scrollbars
    setup_listeners(forms[0])
        
    # Manually insert data in cells for the first time    
    cell = sheet.getCellByPosition(1,1)
    cell.Value = forms[0].getByName("ScrollBar_0").ScrollValue * 10000

    cell = sheet.getCellByPosition(1,2)
    cell.Value = forms[0].getByName("ScrollBar_1").ScrollValue * 0.1

    cell = sheet.getCellByPosition(1,3)
    cell.Value = forms[0].getByName("ScrollBar_2").ScrollValue * 1
    
    cell = sheet.getCellByPosition(1,4)
    cell.Value = forms[0].getByName("ScrollBar_2").ScrollValue * 12    

    # Clear the months column. 
    clear_column(sheet, 0, OFFSET, 500)
    
    setup_chart(doc, sheet)    
    
    recalculate(sheet)

def setup_chart(doc, sheet):
    """ Setup a chart to display the spreadsheet data """
    #Dim Rect As New com.sun.star.awt.Rectangle - in BASIC
    rectangle = uno.createUnoStruct("com.sun.star.awt.Rectangle")

    rectangle.X = 10000
    rectangle.Y = 4000
    rectangle.Width = 14000
    rectangle.Height = 14000
    
    #Dim RangeAddress(0) As New com.sun.star.table.CellRangeAddress - in BASIC
    range_address = []
    range_address.append(uno.createUnoStruct("com.sun.star.table.CellRangeAddress"))
   
    # Range_address parameter of addNewByName expects a sequence of com.sun.star.table.CellRangeAddress     
    range_address[0].Sheet = 0
    range_address[0].StartColumn = 0
    range_address[0].StartRow = OFFSET-1
    range_address[0].EndColumn = 2
    
    cell = sheet.getCellByPosition(1,4)
    total_month = int(cell.Value)
    range_address[0].EndRow = total_month  + (OFFSET -1)
     
    charts = sheet.Charts
    charts.addNewByName("MyChart", rectangle, tuple(range_address), True, True)    
    chart = charts.getByName("MyChart").EmbeddedObject
    
    chart.HasMainTitle = True
    chart.Title.String = "Amortization"
    chart.HasSubTitle = True
    chart.SubTitle.String = "Monthly Interest and Principal"
    
    chart.Diagram = chart.createInstance("com.sun.star.chart.LineDiagram")     
    chart.Diagram.XAxisTitle.String = "Month"
    chart.Diagram.YAxisTitle.String = "Amount"
    chart.Diagram.HasXAxisGrid = True    
    chart.Diagram.XMainGrid.LineColor = 0xC0C0C0 # RGB(192, 192, 192)
    chart.Diagram.HasYAxisGrid = True
    chart.Diagram.YMainGrid.LineColor = 0xC0C0C0 # RGB(192, 192, 192)
    
    chart.Diagram.Wall.FillStyle = SOLID  # from com.sun.star.drawing.FillStyle import SOLID
    chart.Diagram.Wall.FillColor = 0x64A0FF  # RGB(100, 160, 255)
    chart.Diagram.Wall.LineColor = 0x5050FF  # Rgb(80,80,255)
    chart.Diagram.Wall.LineWidth = 80
    chart.Diagram.Wall.LineStyle = LINE_SOLID # com.sun.star.drawing.LineStyle.SOLID 

    chart.HasLegend = True 
    chart.Legend.Alignment = BOTTOM  # com.sun.star.chart.ChartLegendPosition.BOTTOM
    chart.Legend.FillStyle = SOLID  # com.sun.star.drawing.FillStyle.SOLID
    chart.Legend.FillColor = 0xD2D2FF # RGB(210, 210, 255)
    chart.Legend.CharHeight = 10
    
    chart.Area.LineColor = 0x009B00  # Rgb(0,155,0)
    chart.Area.LineWidth = 100
    chart.Area.LineStyle = LINE_SOLID  # com.sun.star.drawing.LineStyle.SOLID as LINE_SOLID
    chart.Area.FillColor = 0xDCFFDC  # Rgb(220,255,220)
    
    # Debugging. Do a dir()
    #debug(chart.Diagram.XAxis)
    # Debugging. Display a property value
    #debug_1(chart.Diagram.XAxis.StepMain)


def setup_listeners(forms_0):
    """ Register the pair of Listeners for each of the 3 x Scrollbars. """

    #Dim oEvent as new com.sun.star.script.ScriptEventDescriptor
    event_descriptor = uno.createUnoStruct( "com.sun.star.script.ScriptEventDescriptor" )    

    # Placed into: While Adjusting. Still doesn#t update when moving the slider.
    # Just updates the Scrollbars spreadsheet cell, does not force recalcualtions.
    listener_interface_name = "com.sun.star.awt.XAdjustmentListener"
    listener_method_name = "adjustmentValueChanged"
    
    PREFIX = "vnd.sun.star.script:"
    FILE = sys.argv[0] + "$"    
    SUFFIX = "?language=Python&location=user" 
    CALLBACK = "cb_scrollbar_adjust"    
    macro_location = PREFIX + FILE + CALLBACK + SUFFIX
    
    #macro_location = "vnd.sun.star.script:calc_uno_python_amortization.py$cb_scrollbar_adjust?language=Python&location=user"    
    # calc_uno_python_amortization.py$cb_scrollbar_mouse_up (user, Python)
    
    # Provide values to the ScriptEventDescriptor parameters
    event_descriptor.AddListenerParam = "" # data to be used if the addListener method needs an additional parameter.  
    event_descriptor.ListenerType = listener_interface_name
    event_descriptor.EventMethod = listener_method_name
    event_descriptor.ScriptType = "Python"  
    event_descriptor.ScriptCode = macro_location

    # Index of the ScrollBar. 0 1 2...
    for i in range(0,3): 
         #oDrawPage.Forms(0).registerScriptEvent(i, event_descriptor)
         forms_0.registerScriptEvent(i, event_descriptor)

    # Place a link against: Mouse button released. This will have a callback to
    # cb_scrollbar_mouse_up routine that perfroms the recalculations
    listener_interface_name = "com.sun.star.awt.XMouseListener"
    listener_method_name = "mouseReleased"

    #macro_location = "vnd.sun.star.script:Module.py$cb_scrollbar_mouse_up?language=Python&location=document"

    CALLBACK = "cb_scrollbar_mouse_up"    
    macro_location = PREFIX + FILE + CALLBACK + SUFFIX
    #macro_location = "vnd.sun.star.script:calc_uno_python_amortization.py$cb_scrollbar_mouse_up?language=Python&location=user" 
    #calc_uno_python_amortization.py$cb_scrollbar_mouse_up (user, Python)
    # The scripting language Python or python is not supported.
    
    event_descriptor.AddListenerParam = "" # data to be used if the addListener method needs an additional parameter.  
    event_descriptor.ListenerType = listener_interface_name
    event_descriptor.EventMethod = listener_method_name
    event_descriptor.ScriptType = "Python"     
    event_descriptor.ScriptCode = macro_location

    # Index of the ScrollBar. 0 1 2...
    for i in range(0,3):
         #oDrawPage.Forms(0).registerScriptEvent(i, event_descriptor)
         forms_0.registerScriptEvent(i, event_descriptor)


def cb_scrollbar_adjust(event):
    """ Scrollbar events while adjusting """
    # as "com.sun.star.awt.AdjustmentEvent"     
    
    #desktop, doc = main_initialize()
    #sheet = doc.Sheets.getByName("Amortization")
    sheet = event.value.Source.Model.Parent.Parent.Parent.Sheets.getByName("Amortization")

    if event.value.Source.Model.Name == "ScrollBar_0":
        # Total loaned
        cell = sheet.getCellByPosition(1, 1)
        cell.Value = event.value.Source.Value * 10000
             
    elif event.value.Source.Model.Name == "ScrollBar_1":
        # Percent per annum
        cell = sheet.getCellByPosition(1, 2)
        cell.Value = event.value.Source.Model.ScrollValue * 0.1
 
    elif event.value.Source.Model.Name == "ScrollBar_2":
        #Term in Years
        cell = sheet.getCellByPosition(1, 3)
        cell.Value = event.value.Source.Model.ScrollValue * 1
        #Term in Months
        cell = sheet.getCellByPosition(1, 4)
        cell.Value = event.value.Source.Model.ScrollValue * 12
        
    else:
        # The cheap and nastie bug reporting system !
        with open("bug.txt", "w") as fout:
            fout.write("Didn't find the name ScrollBar_?\n")


def cb_scrollbar_mouse_up(event):
    """ On Scrollbar Mouse_up then redo the calculations """
    #as com.sun.star.awt.MouseEvent
    sheet = event.value.Source.Model.Parent.Parent.Parent.Sheets.getByName("Amortization")
 
    if event.value.Source.Model.Name == "ScrollBar_0":
        # Total loaned
        cell = sheet.getCellByPosition(1, 1)
        cell.Value = event.value.Source.Value * 10000
             
    elif event.value.Source.Model.Name == "ScrollBar_1":
        # Percent per annum
        cell = sheet.getCellByPosition(1, 2)
        cell.Value = event.value.Source.Model.ScrollValue * 0.1
 
    elif event.value.Source.Model.Name == "ScrollBar_2":
        #Term in Years
        cell = sheet.getCellByPosition(1, 3)
        cell.Value = event.value.Source.Model.ScrollValue * 1
        #Term in Months
        cell = sheet.getCellByPosition(1, 4)
        cell.Value = event.value.Source.Model.ScrollValue * 12
        
    else:
        with open("bug.txt", "w") as fout:
            fout.write("Didn't find the name ScrollBar_?\n")
     
    # Clear the months column. 
    clear_column(sheet, 0, OFFSET, 500)    
    recalculate(sheet)


def clear_column(sheet, col, row, length):
    """ Do this by range. Clear all columns """
    #oSheet = ThisComponent.Sheets.getByName("Amortization")    
    for i in range (0, length):
        for j in range (0, 4):
            cell = sheet.getCellByPosition(j + col, i + row)    
            cell.String = "" 


def recalculate(sheet):
    # Total Monthly Payment = Loan Amount [ i (1+i) ÷ n / ((1+i) ÷ n) - 1) ]
    #sheet = ThisComponent.Sheets.getByName("Amortization")
    cell = sheet.getCellByPosition(1, 1)
    principal = cell.Value
    cell = sheet.getCellByPosition(1, 2)
    interest_percent = cell.Value
    interest = interest_percent * 0.01
    interest_monthly = interest / 12
    
    cell = sheet.getCellByPosition(1, 3)
    years = int(cell.Value)
    cell = sheet.getCellByPosition(1, 4)    
    total_payments = int(cell.Value)
    
    # interest_monthly
    monthly_interest_ratio = (interest_percent * 0.01) / 12

    total_monthly_payment = (( principal * monthly_interest_ratio) / (1- (1+ monthly_interest_ratio) ** - total_payments))
    
    #msgbox total_monthly_payment    
    cell = sheet.getCellByPosition(1, OFFSET-3)
    cell.Value = total_monthly_payment # Already formatted? "${:,.2f}".format(total_monthly_payment) # Currency
     
    #The formula to calculate the monthly principal due on an amortized loan is as follows:
    #Principal Payment = Total Monthly Payment - (Outstanding Loan Balance × (Interest Rate/12 Months))

    total_payment_amount = total_monthly_payment * years * 12
    #msgbox total_payment_amount
    cell = sheet.getCellByPosition(1, OFFSET-4)
    cell.Value = total_payment_amount # Currency    
    # correct for total interest
    total_interest_amount = (total_monthly_payment * years * 12) - principal
    #msgbox total_interest_amount

    # Dim amortization_array(0 to total_payments, 0 to 3) As Single    
    # Write all zeros
    amortization_array = [[0 for i in range(4)] for j in range(total_payments +1)]

    # Enter the month in first column    
    for i in range(0, total_payments + 1):
        amortization_array[i][0] = i
    
    #Insert initial balance, which is the principal
    amortization_array[0][3] = principal

    # Multiply the principal amount by the monthly interest rate: 
    # ($100,000 principal multiplied by 0.005 = $500 month’s interest).
    # You can use the equation: 
    # I=P*r*t, where I=Interest, P=principal, r=rate, and t=time.
    for i in range (1, total_payments +1):
        # previous balance * monthly interest 
        interest_month_amount = amortization_array[i-1][3] * interest_monthly
        repayment_amount = total_monthly_payment - interest_month_amount
        balance = amortization_array[i-1][3] - repayment_amount
        
        amortization_array[i][1] = interest_month_amount
        amortization_array[i][2] = repayment_amount
        amortization_array[i][3] = balance
       
    # Write to spreadsheet. Start at month 1, so Chart will look OK.
    #for i = 0 to total_payments
    for i in range (1, total_payments + 1):
        for j in range (0, 4):
            cell = sheet.getCellByPosition(j, i + OFFSET -1)
            cell.Value = amortization_array[i][j]

    # Charts
    # Update the chart. i.e. Change the row count. Y-Axis values change automatically
    charts = sheet.Charts
    chart = charts.getByIndex(0)
    ranges = chart.getRanges() 
    ranges[0].EndRow = OFFSET + total_payments -1
    chart.setRanges(ranges)


def main():
    """ Main startup """
    print("\nCreating Spreadsheet:", FILE_NAME)
    
    desktop, doc = main_initialize_not_embedded()
    
    # Initial file is title Untitled1. Change to calc_uno_amortization.ods and save
    # This willl overwrite previous files.
    #properties = (PropertyValue('FilterName', 0, 'writer8', 0),) # writer8 - huh???
    properties = (PropertyValue('FilterName', 0, 'calc8', 0),)        
    try:
        # Store As URL if it doesn't exist. Changes local name.
        doc.storeAsURL(FILE_URL, properties)
    except Exception as e:
        #print(e.typeName) # com.sun.star.task.ErrorCodeIOException
        if e.typeName == "com.sun.star.task.ErrorCodeIOException":
            print("Error: ErrorCodeIOException.")
            print("The spreadsheet already exists.")
            doc.dispose()
            sys.exit("Exiting...")
        else:
            print("Error:", e.typeName)
            sys.exit("Exiting...")

    # change name from Sheet1 to Amortization.
    sheet = doc.Sheets.getByName("Sheet1")
    sheet.setName("Amortization")

    setup_cell_format(doc, sheet)    

    setup_slider_initial(doc, sheet)
    
    # Completed creating sheet, so turn off design mode.
    doc.getCurrentController().setFormDesignMode(False)


# Lists the scripts, that shall be visible inside OOo. Can be omitted, if
# all functions shall be visible. Make callbacks to form controls visible.
g_exportedScripts = (cb_scrollbar_adjust, cb_scrollbar_mouse_up, main,)
    
    
if __name__=="__main__":

    main()
