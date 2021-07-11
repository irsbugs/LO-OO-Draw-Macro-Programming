# LibreOffice or OpenOffice Draw Application. 

# Drawing using embedded BASIC programming.

The LibreOffice/OpenOffice (LO/OO) suite of applications all support the BASIC programming language being embedded in documents. The Draw application contains shapes, for example rectangle and ellipse, that may be programatically drawn onto the document page being created.

Draw uses a resolution down to 1/100th of a millimeter. Thus an A4 sized page which is 210mm x 297mm is comprised of 21000 x 29700 points. While diagrams may be manually drawn, greater accuracy is attained by using BASIC to draw the shapes.

LO/OO also support adding form controls to a Draw document. For example, push-button controls may be added which can be used to re-draw the diagrams in different ways.

Attached is the Draw document:

* floor_plan.odg

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
Tools --> Macros --> Edit Macros --> floor_plan.odg --> Standaard --> Module1. 
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

## Examples of BASIC code.

* Guide to programming in BASIC.

A Guide to BASIC programming for LO is available online: 
https://help.libreoffice.org/latest/en-GB/text/sbasic/shared/main0601.html?DbPAR=BASIC

This includes an alphabetic list of command which starts here: 
https://help.libreoffice.org/latest/en-GB/text/sbasic/shared/03080601.html?DbPAR=BASIC


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
```
0, MB_OK - OK button
1, MB_OKCANCEL - OK and Cancel button
2, MB_ABORTRETRYIGNORE - Abort, Retry, and Ignore buttons
3, MB_YESNOCANCEL - Yes, No, and Cancel buttons
4, MB_YESNO - Yes and No buttons
5, MB_RETRYCANCEL - Retry and Cancel buttons

0, MB_DEFBUTTON1 - First button is default value
256, MB_DEFBUTTON2 - Second button is default value
512, MB_DEFBUTTON3 - Third button is default value				

16, MB_ICONSTOP - Stop sign
32, MB_ICONQUESTION - Question mark
48, MB_ICONEXCLAMATION - Exclamation point
64, MB_ICONINFORMATION - Tip icon

1, IDOK - Ok
2, IDCANCEL - Cancel
3, IDABORT - Abort
4, IDRETRY - Retry
5 - Ignore
6, IDYES - Yes
7, IDNO - No
```	
	 
