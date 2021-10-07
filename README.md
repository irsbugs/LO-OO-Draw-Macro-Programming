# LibreOffice or OpenOffice Draw Application. 

Drawing using macro scripts:
* Embedded BASIC
* Python with Universal Network Objects *UNO*.
* Embedded Python

# Drawing using embedded BASIC programming.

The LibreOffice/OpenOffice (LO/OO) suite of applications all support the BASIC programming language being embedded in documents. The Draw application contains shapes, for example rectangle and ellipse, that may be programatically drawn onto the document page being created.

Draw uses a resolution down to 1/100th of a millimeter. Thus an A4 sized page which is 210mm x 297mm is comprised of 21000 x 29700 points. While diagrams may be manually drawn, greater accuracy is attained by using BASIC to draw the shapes.

LO/OO also support adding form controls to a Draw document. For example, push-button controls may be added which can be used to re-draw the diagrams in different ways.

Attached is the Draw document:

* draw_embedded_basic_plan.odg

It includes BASIC script which Draws every part of the diagram. It shows a floor area of a house and allows three alternative layouts of the piles to be reviewed by clicking on buttons.

The shapes used to draw the *floor_plan.odg* diagram are:
* Rectangle
* Ellipse
* Line
* Measure line
* Text boxes
* Polygon

To access the BASIC script click on: 
```
Tools --> Macros --> Edit Macros --> floor_plan.odg --> Standard --> Module1. 
```
An Integrated Developemnt Environment window then provides a means to edit and run the BASIC scripting.

The BASIC script that is embedded in *floor_plan.odg* has been attached as the file *floor_plan.bas*. This file may be reviewed and copied to retrieve sections of code to use in other documents.

## Known Issues

1. There is a boolean feature named *Design Mode*. It may be manually toggled by clicking:
```
Tools --> Forms --> Design Mode
```
Although this toggling of *Design Mode* may be performed programatically, this does not appear to work. Thus upon completion of starting the BASIC script it may be necessary to manually toggle *Design Mode* a few times in order for the control push-buttons to commence operation.

2. By default the Layers provided which have visual controls are: 

* layout ~ layer Identification = 0 ~ Displayed as: Layout
* controls ~ layer Identification = 3 ~ Displayed as: Controls
* measurelines ~ layer Identification = 4 ~ Displayed as: Dimension Lines

Additionally there are two layers automatically provided for which there are no visual controls:
* background ~ layer Identification = 1 ~ Not displayed
* backgroundobjects ~ layer identification = 2 ~ Not displayed

The BASIC script adds the following three layers:
* Border ~ layer Identification = 5 ~ Displayed as: Borders
* Grid ~ layer identification = 6 ~ Displayed as Grid
* Piles ~ layer identification = 7 ~ Disploayed as Piles

The controlable layers from the above should be able to be modified to enable and disable the visibility of the elements in the layer. This feature seem to be unstable.

3. There does not appear to be a way to programmatically set the overall scale of a diagram. This is required if Dimension Lines are used in the diagram. To manually set the scale of 1:80 for the *floor_plan.odg* preform the following:
```
Tools--> Options --> LibreOffide Draw--> General--> 
```
then set:
```
Drawing Scale: 1:80
Unit of measure: Meter.
```
## References

* Guides to programming in BASIC.

A Guide to BASIC programming for LO is available online: 
https://help.libreoffice.org/latest/en-GB/text/sbasic/shared/main0601.html?DbPAR=BASIC

This includes an alphabetic list of command which starts here: 
https://help.libreoffice.org/latest/en-GB/text/sbasic/shared/03080601.html?DbPAR=BASIC



## Examples of BASIC code.

The following is an overview with examples from the BASIC script used for *floor_plan.odg*.

* A comment may be added by prefixing the comment with REM or a single quote mark.

* Dimensioning of a variable may be performed in the module outside of the subroutines and functions. This is normally in the preamble of a module and variables defined as *Public* are avaialble to all subroutines and functions in the module. For example:
```
Public oDoc as object
```

* Constants may be defined in the module outside of subroutines and functions and they become available to all routines. For examples:
```
Const PILE_X_SIZE as integer = 250
Const PILE_Y_SIZE as integer = 250
```

* A jump over code, in order to bypass testing code, may be performed with the *goto* command:
```
	goto skip_layer_ident
	' Debugging code to check which Layers exist and their identifications.
	dim i as integer
	for i = 0 to oLM.Count -1
		oLayer = oLM.getByIndex(i)
		msgbox "Layer Name: " + oLayer.Name + chr(13) + "Layer ID: " + cstr(i) 
	next i
	exit sub
	
	skip_layer_ident:
	main
```

* A Dimension line is provided by the following subroutine:
Call: ruler(3000, 4000, 12000, 0, 1500, False)
i.e.:
Start at point X=3000, Y=4000, Horizontal Width = 12000, Vertical Height = 0, Offset from the line of 1500 points, Place in default position.

```
sub ruler(X as long, Y as long, W as long, H as long, optional MDL as integer, _
	optional MBRE as boolean)
	' Routine to draw measurment lines as layer 4 on a page
	' MDL = MeasureLineDistance. Offset of line from the two points
	if ismissing(MDL) then
		MDL = 1000
	end if	
	' MBRE = MeasureBelowReferenceEdge. Above or below the two points.
	if ismissing(MBRE) then
		MBRE = False
	end if	
	Point.x = X
	Point.y = Y
	Size.Width = W
	Size.Height = H		

	MeasureShape = oDoc.createInstance("com.sun.star.drawing.MeasureShape")
	MeasureShape.Size = Size
	MeasureShape.Position = Point	
	oPage.add(MeasureShape)
	' Changes to font must be after adding to the page
	MeasureShape.LayerID = 4		
	MeasureShape.LineColor = 0
	MeasureShape.LineWidth = 5
	MeasureShape.MeasureLineDistance = MDL
	MeasureShape.MeasureBelowReferenceEdge = MBRE
	MeasureShape.CharWeight = com.sun.star.awt.FontWeight.NORMAL 'BOLD	
	MeasureShape.CharFontName = "FreeSans" '"Ubuntu Mono"
	MeasureShape.CharHeight = 12				
end sub
```
The subroutine is passed the start point for the line, the X and Y points. The width and height, W and H values, to determine the end point of the dimension being measured. The offset in 1/100 of a mm of the dimension line from the points being measured. The direction of the offset, above or below the line being measured.

* Drawing a set of lines using PolPolygonShape. The Polygon is simple compared to drawing individual lines.
```
sub border_line
	' Draw a border at 600 on a sheet of Landscape A4. 
	' Polygon more simple than drawing individual lines.
	Dim oDoc as object
	Dim oPage as object	
	Dim Square(3) As New com.sun.star.awt.Point

	oDoc = ThisComponent
	oPage = oDoc.DrawPages(0)
	
	' Clear the Page of all elements...
	for i = oPage.getCount - 1 to 0 step -1
		oPage.Remove(oPage.getByIndex(i))
	next i

	wait 2000

	PolyPolygonShape = oDoc.createInstance("com.sun.star.drawing.PolyPolygonShape")
	'PolyPolygonShape.LayerID = 5		
	PolyPolygonShape.LineColor = 0
	PolyPolygonShape.LineWidth = 10
	PolyPolygonShape.FillTransparence = 100
		 
	oPage.add(PolyPolygonShape) 
	' Page.add must take place before the coordinates are set
		 
	Square(0).x = 600
	Square(1).x = 29100
	Square(2).x = 29100
	Square(3).x = 600
	Square(0).y = 600
	Square(1).y = 600
	Square(2).y = 20400
	Square(3).y = 20400

	PolyPolygonShape.PolyPolygon = Array(Square())
end sub
```

* A message box dialog
```
	' Using msgbox...
	s = "The following control will be removed: " + chr(13) 
	if msgbox (s + "Layer ID: " + cstr(element.LayerID) + chr(13) + _
		   "Layer Name: " + element.LayerName + chr(13) + _
		   "Control Name: " + element.control.name + chr(13) + _
		   "Control Label: " + element.control.label, _
		   MB_YESNO + MB_DEFBUTTON1 + MB_ICONQUESTION, _
		   "Remove Control") = IDYES then
		   
		do x...
	else
		do y...
	end if
```	
Constants for msgbox:

* Button combination:
```
0, MB_OK - OK button
1, MB_OKCANCEL - OK and Cancel button
2, MB_ABORTRETRYIGNORE - Abort, Retry, and Ignore buttons
3, MB_YESNOCANCEL - Yes, No, and Cancel buttons
4, MB_YESNO - Yes and No buttons
5, MB_RETRYCANCEL - Retry and Cancel buttons
```
* Button highlighted to be accepted as the default:
```
0, MB_DEFBUTTON1 - First button is default value
256, MB_DEFBUTTON2 - Second button is default value
512, MB_DEFBUTTON3 - Third button is default value				
```
* Icon to be displayed
```
16, MB_ICONSTOP - Stop sign
32, MB_ICONQUESTION - Question mark
48, MB_ICONEXCLAMATION - Exclamation point
64, MB_ICONINFORMATION - Tip icon
```
* Returned values depending on button clicked:
```
1, IDOK - Ok
2, IDCANCEL - Cancel
3, IDABORT - Abort
4, IDRETRY - Retry
5, IDIGNORE - Ignore
6, IDYES - Yes
7, IDNO - No
```	

# Drawing using Python programming and UNO.

The LibreOffice/OpenOffice (LO/OO) suite of applications support a Python programming script communicating with LO/OO via the 
Universal Network Objects *UNO*. The Draw application contains shapes, for example rectangle and ellipse, that may be 
programatically drawn onto the document page being created. Also Form Controls, such as buttons, may be dynamically added
to the page that is drawn.

The attached Python program provides a demonstration:

* draw_uno_plan.py

## Installation

On Linux systems, for example Ubuntu Mate 20.04.2 with LibreOffice 6.4.7.2, this file is located at:
```
 ~/.config/libreoffice/4/user/Scripts/python/draw_uno_plan.py
```

The system must include Python3 and the Python module uno. Install this with the command:
```
$ sudo apt install python3-uno
```

## Launching draw_uno_plan.py:

* Open two console terminal windows.

* In one terminal window enter the command:
```
 $ libreoffice --draw --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"
```

* In the other terminal window enter the command:
```
$ python ~/.config/libreoffice/4/user/Scripts/python/draw_uno_plan.py
```

The LibreOffice Draw application will be launched and the program will add the shapes to create the 
drawing. Also added are three Form Control Buttons. After this the Python program terminates.

The drawing is 12m x 10m floor plan with a one meter reference grid. Clicking on the three buttons
allows modelling of different piles layouts for the floor. Thus, although the Python program has terminated, 
the buttons allow running a series of functions within the program.

The program may also be run from the LibreOffice Menu bar with:

Tools--> Macro--> Run Macros --> Library: My Macros --> draw_uno_plan --> Macro Name: main --> Run

## Notes and Reference links for writing Python programs for LO/OO

1.  Be aware of case sensitivity.

2.  References: 

    http://christopher5106.github.io/office/2015/12/06/openoffice-libreoffice-automate-your-office-tasks-with-python-macros.html
    https://wiki.documentfoundation.org/Macros/Python_Guide/Introduction
    https://wiki.documentfoundation.org/Macros/Python_Design_Guide
    https://www.scribd.com/document/75405001/OpenOffice-org-Developer-s-Guide-Professional-UNO    
    https://wiki.openoffice.org/wiki/Python/Transfer_from_Basic_to_Python  
    https://wiki.openoffice.org/wiki/Python/Transfer_from_Basic_to_Python#Script_Context
    https://forum.openoffice.org/en/forum/viewtopic.php?f=20&t=66707&p=296638&hilit=CreateButton#p296638
     
3.  This code uses uno module rather than XSCRIPTCONTEXT. See...
    https://wiki.openoffice.org/wiki/PyUNO_samples - TableSample.py

4.  Change to be a file. required by "AssignAction()  ScriptEventDescriptor.ScriptCode"
    Typical BASIC...
    ```
    sScriptURL = "vnd.sun.star.script:Standard.Module1.ButtonPushEvent?language=Basic&location=document"
    ```
    Some other link...
    ```
    sScriptURL = "vnd.sun.star.script:ScriptBindingLibrary.MacroEditor?location=application"
    ```
    above equates to:`ScriptBindingLibrary.MacroEditor (application, )`
    
    For program: ~/.config/libreoffice/4/user/Scripts/python/draw_uno_plan.py:
    ```
    Function: button_push_event(button):
    sScriptURL = "vnd.sun.star.script:draw_uno_plan.py$button_push_event?language=Python&location=user"
    ```
    This equates to:
    ```
    Events, Execute Action: draw_uno_plan.py$button_push_event (user, Python)
    ```
    
5.  aEvent = uno.createUnoStruct("com.sun.star.script.ScriptEventDescriptor")
    aEvent has these structures...
    ```
    ListenerType:	listener type as string, same as listener-XIdlClass.getName().  
    EventMethod:	event method as string.  
    AddListenerParam:	data to be used if the addListener method needs an additional parameter.  
    ScriptType:	    type of the script language as string; for example, "Basic" or "StarScript".  
    ScriptCode:	    script code as string (the code has to correspond with the language defined by ScriptType).          
    ``` 
   
6.  The BASIC msgbox does not work with python. A messagebox function omsgbox() 
    is available. It only displays strings. Place anywhere to help debug code. E.g.
    ```
    dir_list = dir(uno)
    omsgbox((", ").join(dir_list), "Python dir() Listing")
    ```
7.  Program control of "Design Mode" is suspect. May need to be toggled a few times.
    ```
    Example for BASIC
    Global b as Boolean
    Sub toggleFormDesignMode()
        c = ThisComponent.getCurrentController()
        c.setFormDesignMode(b)
        b = Not b
    End Sub
    ```

## Screenshot

<img src="https://github.com/irsbugs/LO-OO-Draw-Macro-Programming/blob/main/screenshot_floor_plan.png">


# Python embeded in a LO/OO document

BASIC is the default language that LO/OO support for being embedded into a document. The support includes a BASIC IDE. Examination of an ODF file with embedded BASIC reveals a directory structure of Basic/Standard/Module1.xml, where Module1.xml is the default name for the module that contains the BASIC code, stored in XML format. The META-INF/manifest.xml file contains details of the Basic files.

When embedding Python scripts into a LO/OO document the process is described here...

https://wiki.documentfoundation.org/Macros/Python_Guide/Introduction#Installation


However, this process is simplified if an extension tool is added to LO/OO. This is the Alternative Python Script Organizer (APSO). It may be downloaded from here...

https://extensions.libreoffice.org/en/extensions/show/apso-alternative-script-organizer-for-python

Once it has been installed then click on Tools--> Extension Manager... and click on APSO so it is highlighted. Then click on Options. Set the Editor to the IDE normally used to develop Python. For example if you use Geany on a Linux system, enter /usr/bin/geany

APSO is accessed by typing Shift+Alt+F11 or clicking on Tools--> Macros --> Organize python scripts. A new Module may then be created, and when the module is edited the IDE will launch so that entering python code may begin.

Upon saving the IDE editing session, closing APSO, and saving and closing the LO/OO document, then upon the next time the document is launched is will have the python code available for execution.

With embedded python code the document has the code stored at /Scripts/python/Module.py. The APSO tools embeds the python script and performs the modifications to the META-INF/manifest.xml.

Attached are two examples of documents with embedded code. They are both Writer documents that contains two push-buttons labelled "Message" and "Clear". One document has embedded BASIC code, while the other has embedded Python. when the "message" button is pushed it displays in the document a "Hello World" message followed by listing the contents of the embedded program. These programs are:

* writer_basic_example.odt
* writer_python_example.odt

# Drawing using embedded Python programming.

Python scripts may be embedded in Draw documents. 

The following document was the *draw_embedded_basic_plan.odg*, which then had its BASIC code re-written in Python and made use of UNO to be *draw_uno_plan.py*. This code was then embedded and modified with the aid of APSO to run as embedded Python code in the document:

* draw_python_embedded_plan.odg

Two buttons are added at the top of the document. They are labelled "Start" and "Clear". Clicking *Clear* will remove all elements from the drawing. It also remove three layers that get added. Clicking "Start" will run the Python code to draw a floor plan and add layers.

# Calc using embedded Python programming

The following document is a Calc spreadsheet. On *Sheet1* are the push buttons *Clear* and *Create*. These buttons create and remove a sheet labelled *Amortization*. On this sheet three ScrollBars are provided so that a loan may be modelled and the amortization of the loan is made available as data and in a chart. The creating and modelling of the data is performed by Python code, which is embedded in the Calc document. 

* calc_embedded_python_amortization.ods

# Python program using UNO to create Calc document.

The following python program will launch a Calc document and create a sheet with amortization data that may be modelled with scrollbars. 

* calc_uno_python_amortization.py

## Screenshot

<img src="https://github.com/irsbugs/LO-OO-Draw-Macro-Programming/blob/main/screenshot_amortization.png">


# Python Msgbox

The BASIC language provides a Message Box *msgbox* function, while Python does not have this feature.

APSO comes with a small library **apso_utils** that contains some helpers for debugging purpose during macro writing. These helpers are four functions named *msgbox*, *xray*, *mri* and *console* and can be imported by a classic import statement like **from apso_utils import msgbox**

The following writer document demonstrates a msgbox with python.

* writer_python_msgbox.odt



