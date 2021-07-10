REM  *****  BASIC  *****
'
' Filename: floor_plan.odg.

' Objective:
' Provide a plan diagram of the proposed house using the LibreOffice/OpenOffice
' draw application.
' Use embedded BASIC to Draw the plan and provide modelling.
' Utilize the Layer Manager for different aspects of the plan to be on different
' layers.
'
' Author: Ian Stewart
' Repository: github.com/irsbugs/maison
' Date: Jun 2021
'
' Scale: 
' 1cm = 80cm, ie. 1.25cm = 1m.
' Grid points: 1000 = 1cm.
' Page size A4 Landscape = 29700 x 21000 grid points
' Set scale to 1:80:
' Tools--> Options --> LibreOffide Draw--> General--> Drawing Scale 1:80
' Unit of measure: Meter.

' Ruler  1         2         3         4         5         6         7         8
'2345678901234567890123456789012345678901234567890123456789012345678901234567890

' Variables dimensioned for all subroutines  
Public oDoc as object
Public oPage as object
Public oLM as object
Public Point As New com.sun.star.awt.Point
Public Size As New com.sun.star.awt.Size

' Constants
Const LABEL as string = "A4 Landscape. Scale 1:80"
'Dim M1 as integer ' 1 meter = 1250 grid points
Const M1 as integer = 1250 

' Pile dimensions are 200mm x 200mm. 100mm = 125 points	
Const PILE_X_SIZE as integer = 250
Const PILE_Y_SIZE as integer = 250


' Notes:
' The boolean "DesignMode" is used when adding Form Controls (i.e. widgets).
' It is not dynamic, and may need to be manually toggled on/off after adding a
' Form Control. Thus Form Controls are only added once and not deleted. 

sub start_normal
	' This start does not remove existing controls added by start_initial.
	global_instantiation
	
	'dim i as integer
	'for i = 0 to oLM.Count -1
	'	oLayer = oLM.getByIndex(i)
	'	msgbox "Layer Name: " + oLayer.Name + chr(13) + "Layer ID: " + cstr(i) 
	'next i
	
	main
	
end sub

sub start_initial()
	' Remove and add the buttons. This is likely to leave the Draw page in 
	' "DesignMode". Which will require being manually disabled by:
	' Tools --> Forms --> Design Mode (uncheck)
	global_instantiation
    ' Tools --> Form Controls --> "automatic contorl Focus" -huh?

	dim element as object

	' remove all button controls. These exist on Layer 3
	for i = oPage.getCount - 1 to 0 step -1
		' Order is reversed to allow to ensure index pointer doesn't change.
		element = oPage.getByIndex(i)
		'msgbox element.name + " " + element.label + " " + _
		'		Cstr(i) + " " + Cstr(element.LayerID)
		if element.LayerID = 3 then
			' Layer 3 contains the control object(s)
			s = "The following control will be removed: " + chr(13) 
			if msgbox (s + "Layer ID: " + cstr(element.LayerID) + chr(13) + _
				   "Layer Name: " + element.LayerName + chr(13) + _
				   "Control Name: " + element.control.name + chr(13) + _
				   "Control Label: " + element.control.label, _
				   MB_OK + MB_DEFBUTTON1 + MB_ICONQUESTION, _
				   "Remove Control") = IDOK then
				   
				'msgbox "to be removed"
				oPage.Remove(element)				
			else:
				' If msgbox forced to close
				msgbox "Not Removed. Exiting..."
				exit sub
			end if				   
		end if
	next i
	
	' All controls have been removed. So add them...
	add_control

	' These have no effect. Need to manually toggle.
	oDoc.ApplyFormDesignMode = True
	oDoc.ApplyFormDesignMode = False	
	
	main
	
	msgbox ("Manually toggle Design Mode to activate buttons." + chr(13) + _
			"Tools --> Forms --> Check then Uncheck 'Design Mode'", _
			MB_OK + MB_DEFBUTTON1 + MB_ICONEXCLAMATION, _
			"Disable Design Mode")	
end sub

sub add_control
	' Add the control buttons
	Dim aName As Variant
	Dim aLabel As Variant	
	aName = Array("B0", "B1", "B2")
	aLabel = Array("6m x 5m", "4m x 5m", "4m x 3.33m")
	
	for i = 0 to 2
		create_button(aName(i), aLabel(i), i)
	next i		
end sub

sub global_instantiation
	' Global Constants used in subroutines Dimmed in preamble.
	'LABEL =  "A4 Landscape. Scale 1:80"
	'M1 = 1250

	' instantiation.
	oDoc = ThisComponent
	oPage = oDoc.DrawPages(0)
	oLM = oDoc.getLayerManager()
	
	'oDoc  SbxBOOL ApplyFormDesignMode <-- Fails?
end sub
	
Sub Main
	' Main subroutine to create the initial Drawing. After this Control Buttons
	' allow modelling of the drawing.
	global_instantiation

	clear_elements ' Except for Control Buttons
	
	a4_setup
	
	' If the borders layer doesn't exist as index 5 then add it.
	add_border_layer
	
	add_grid_layer
	
	border_line
	
	border_text_field
	
	'Msgbox oPage.Width ' 29700 - 1200 = 28500 Max: 22.8 meters
	'Msgbox oPage.Height '21000 - 1200 = 19800 Max: 15.84 meters

	compass
		
	grid
	
	grid_supplement

	' Measurment lines:	
	' House horizontal. Pile centers
	ruler(3000, 4000, M1 *12, 0, 1500, False)
	' Grid square of 1 meter 
	ruler(3000, 4000, M1, 0, 800, False)
	' House vertical. Pile centers
	ruler( 3000, 4000, 0, M1 * 10, 800, True)
	' Overall house horizontal
	ruler(3000, 4000, M1*19, 0, 2200, False)
	' Additional Grid on RHS
	ruler(3000+M1*19, 4000, 0, M1*4, 1000, False)
	' House vertical outside of pile
	ruler(3000, 4000-125, 0, (M1*10)+(125*2), 1600, True)		
	' House horizontal outside of pile
	ruler(3000-125, 4000 + M1*10 + 125, (M1*12)+(125*2), 0, 1500, True)
	
	add_pile_layer

end sub

sub create_button(sName As String, sLabel As String, index as integer)
	' Dynamically create a button.
	' https://forum.openoffice.org/en/forum/viewtopic.php?f=20&t=66707&p=296638&hilit=CreateButton#p296638
	' Requires routines: AssignAction, AddNewButton, GetIndex, ButtonPushEvent
	
	sScriptURL = "vnd.sun.star.script:Standard.Module1.ButtonPushEvent?language=Basic&location=document"
	oButtonModel = AddNewButton(sName, sLabel, oDoc, oPage, index)  
	oForm = oPage.getForms().getByIndex(0)
	' find index inside the form container
	nIndex = GetIndex(oButtonModel, oForm)
	AssignAction(nIndex, sScriptURL, oForm)

	' ApplyFormDesignMode fails to change the Design mode.
	oDoc.ApplyFormDesignMode = False
end sub

Sub AssignAction(nIndex As Integer, sScriptURL As String, oForm As Object)
	' assign sScriptURL event as css.awt.XActionListener::actionPerformed.
	' event is assigned to the control described by the nIndex in the 
	' oForm container
	
	aEvent = CreateUnoStruct("com.sun.star.script.ScriptEventDescriptor")
	with aEvent
		.AddListenerParam = ""
		.EventMethod = "actionPerformed"
		.ListenerType = "XActionListener"
		.ScriptCode = sScriptURL
		.ScriptType = "Script"
	end with
	
	oForm.registerScriptEvent(nIndex, aEvent)

end sub

function AddNewButton(sName As String, sLabel As String, oDoc As Object, _
		oPage As Object, index as Integer) As Object
	oControlShape = oDoc.createInstance("com.sun.star.drawing.ControlShape")
	Point.X = 1000 + (3000*index)
	Point.Y = 19700
	Size.Width = 2500
	Size.Height = 600
	oControlShape.setPosition(Point)
	oControlShape.setSize(Size)
	
	oButtonModel = CreateUnoService("com.sun.star.form.component.CommandButton")
	oButtonModel.Name = sName
	oButtonModel.Label = sLabel
	   
	oControlShape.setControl(oButtonModel)
	oPage.add(oControlShape)
	' Layer 3 is the Controls Layer for Form widgets.
	oControlShape.LayerID = 3
	'oControlShape.LayerName = "Controls"
	AddNewButton = oButtonModel

end function

function GetIndex(oControl As Object, oForm As Object) As Integer
	Dim nIndex As Integer
	nIndex = -1
	For i = 0 To oForm.getCount() - 1 step 1
		If EqualUnoObjects(oControl, oForm.getByIndex(i)) Then
			nIndex = i
			Exit For
		End If
  	Next
  	GetIndex = nIndex
end function

sub ButtonPushEvent(ev as com.sun.star.awt.ActionEvent)
	' All buttons run this sub-routine when clicked.
	clear_pile
	
	select case ev.source.Model.Name
		case "B0"
			add_pile_0
		
		case "B1"
			add_pile_1
				
		case "B2"
			add_pile_2	
	end select	
end sub

sub compass
	' Draw an arrow pointing in North direction. Put it in a circle.
	' TODO?: FillBitmapRectanglePoint  .drawing.RectanglePoint MIDDLE_MIDDLE
	dim x as integer
	dim y as integer
	x = 28000
	y = 18000

	Point.x = x
	Point.y = y
	Size.Width = 0 ' 1 meters @ 1:80 4/5th of 1250 is 1000/100 = 1 meter 
	Size.Height = 1000
	LineShape = oDoc.createInstance("com.sun.star.drawing.LineShape")
	LineShape.Size = Size
	LineShape.Position = Point
	LineShape.LayerID = 6		
	LineShape.LineColor = RGB(0,0,255)
	LineShape.LineWidth = 50

	oPage.add(LineShape)	
	
	' Must add arrow after adding object to page
		
	LineShape.LineStartWidth = 200
	LineShape.LineStartName = "Arrow"

	'LineShape.LineEndWidth = 200
	'LineShape.LineEndName = "Circle"
	LineShape.RotateAngle = 3000 ' 3000 = 30 degrees anti clockwise 
	'							   from horizontal east to west	

	Point.x = x - 500
	Point.y = y
	Size.Width = 1000 ' 1 meters @ 1:80 4/5th of 1250 is 1000/100 = 1 meter 
	Size.Height = 1000	
	
	EllipseShape = oDoc.createInstance("com.sun.star.drawing.EllipseShape")
	EllipseShape.Size = Size
	EllipseShape.Position = Point	
	oPage.add(EllipseShape)
	EllipseShape.LineColor = 0
	EllipseShape.FillColor =  rgb(0,255,0)
	EllipseShape.LineWidth = 5
	EllipseShape.FillTransparence = 80

	
	Point.x = x - 350
	Point.y = y + 200
	Size.Width = 500 ' 1 meters @ 1:80 4/5th of 1250 is 1000/100 = 1 meter 
	Size.Height = 500
	
	TextShape = oDoc.createInstance("com.sun.star.drawing.TextShape")
	TextShape.Size = Size
	TextShape.Position = Point

	oPage.add(TextShape)
	TextShape.String = "N"
	TextShape.CharColor = RGB(255,0,0)	
	TextShape.CharFontName = "FreeSans" '"Ubuntu Mono"
	TextShape.CharHeight = 12
	'msgbox TextShape.TextHorizontalAdjust	
	'TextShape.TextHorizontalAdjust = 1
	'msgbox TextShape.ParaLeftMargin
	
end sub

sub add_pile_0
	' Add a grid of piles 6 meters apart horizontal and 5m apart vertical
	' 9 piles. Calls Pile subroutine
	' 1m = M1 = 1250
	for i = 0 to 2
		for j = 0 to 2 
			Pile(i*6*M1 + 2900, j*5*M1 + 3900)	
		next j
	next i	
	
	'element = oPage.getByName("Page Label")
	for i = 0 to oPage.getCount - 1
		element = oPage.getByIndex(i)	
		if element.name = "Page Label" then
			element.string = "6m x 5m. 9 Piles. " + LABEL
			exit for
		end if
	next i			
end sub

sub add_pile_1
	' Add a grid of piles 4 meters apart horizontal and 5m apart vertical
	' 12 piles. Calls Pile subroutine
	' 1m = M1 = 1250
	for i = 0 to 3
		for j = 0 to 2 
			Pile(i*4*M1 + 2900, j*5*M1 + 3900)	
		next j
	next i
	'element = oPage.getByName("Page Label")
	for i = 0 to oPage.getCount - 1
		element = oPage.getByIndex(i)	
		if element.name = "Page Label" then
			element.string = "4m x 5m. 12 Piles. " + LABEL
			exit for
		end if
	next i	
	
end sub

sub add_pile_2
	' Add a grid of piles 4 meters apart horizontal and 3.33m apart vertical
	' 16 piles Calls Pile subroutine
	' 1m = M1 = 1250
	for i = 0 to 3
		for j = 0 to 3
			Pile(i*4*M1 + 2900, j*3.333*M1 + 3900)	
		next j
	next i
	'element = oPage.getByName("Page Label")
	for i = 0 to oPage.getCount - 1
		element = oPage.getByIndex(i)	
		if element.name = "Page Label" then
			element.string = "4m x 3.33m. 16 Piles. " + LABEL
			exit for
		end if
	next i	
end sub



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

sub grid
	' 1 meter grid over the floor area
	dim i, j
	dim x as integer
	dim y as integer
	
	x = 3000
	y = 4000

	for i = 0 to 12
		Point.x = x + (i * M1)
		Point.y = y
		Size.Width = 0 ' 1 meters @ 1:80 4/5th of 1250 is 1000/100 = 1 meter 
		Size.Height = 12500
		LineShape = oDoc.createInstance("com.sun.star.drawing.LineShape")
		LineShape.Size = Size
		LineShape.Position = Point
		LineShape.LayerID = 6		
		LineShape.LineColor = 0
		LineShape.LineWidth = 10
		
		LineShape.LineStyle = com.sun.star.drawing.LineStyle.DASH		
		' Create a new DashLine object
		oLineDash = LineShape.LineDash
		' Set the values of the object
		oLinedash.Style = 1 ' 0 = Dashes 1 & 4 = dots 2 = None huh?
		oLineDash.Dots = 2 
		oLineDash.DotLen = 1
		oLineDash.Dashes = 1
		oLineDash.DashLen = 100
		oLineDash.Distance = 50
		
		' Accept the new parameters
		LineShape.LineDash = oLineDash
		
		' Check new settings
		dim s
		s = "LineShape.LineStyle: " + Cstr(LineShape.LineStyle)
		s = s + " Style: " + Cstr(LineShape.LineDash.Style)
		s = s + " Dots: " + Cstr(LineShape.LineDash.Dots)
		s = s + " DotLen: " + Cstr(LineShape.LineDash.DotLen)
		s = s  + " Dashes: " + Cstr(LineShape.LineDash.Dashes)		
		s = s  + " DashLen: " + Cstr(LineShape.LineDash.DashLen)
		s = s  + " Distance: " + Cstr(LineShape.LineDash.Distance)		
		'msgbox s 
				
		oPage.add(LineShape)				
	next i
	
	for i = 0 to 10
		Point.x = x
		Point.y = y + (i * M1)
		Size.Width = 15000 ' 1 meters @ 1:80 4/5th of 1250 is 1000/100 = 1 meter 
		Size.Height = 0
		LineShape = oDoc.createInstance("com.sun.star.drawing.LineShape")
		LineShape.Size = Size
		LineShape.Position = Point
		LineShape.LayerID = 6		
		LineShape.LineColor = 0
		LineShape.LineWidth = 10		

		LineShape.LineStyle = com.sun.star.drawing.LineStyle.DASH		
		' Create a new DashLine object
		oLineDash = LineShape.LineDash
		' Set the values of the object
		oLinedash.Style = 1 ' 0 = Dashes 1 & 4 = dots 2 = None huh?
		oLineDash.Dots = 2 
		oLineDash.DotLen = 1
		oLineDash.Dashes = 1
		oLineDash.DashLen = 100
		oLineDash.Distance = 50
		
		' Accept the new parameters
		LineShape.LineDash = oLineDash		
			
		oPage.add(LineShape)	
	next i
end sub

sub	grid_supplement
	' Add a supplement to the grid to the right of 7m x 4m
	' 1 meter grid over the floor area
	dim i, j
	dim x as integer
	dim y as integer
	
	x = 18000
	y = 4000

	for i = 1 to 7
		Point.x = x + (i * M1)
		Point.y = y
		Size.Width = 0 ' 1 meters @ 1:80 4/5th of 1250 is 1000/100 = 1 meter 
		Size.Height = 5000
		LineShape = oDoc.createInstance("com.sun.star.drawing.LineShape")
		LineShape.Size = Size
		LineShape.Position = Point
		LineShape.LayerID = 6		
		LineShape.LineColor = 0
		LineShape.LineWidth = 10
		
		LineShape.LineStyle = com.sun.star.drawing.LineStyle.DASH		
		' Create a new DashLine object
		oLineDash = LineShape.LineDash
		' Set the values of the object
		oLinedash.Style = 1 ' 0 = Dashes 1 & 4 = dots 2 = None huh?
		oLineDash.Dots = 2 
		oLineDash.DotLen = 1
		oLineDash.Dashes = 1
		oLineDash.DashLen = 100
		oLineDash.Distance = 50
		
		' Accept the new parameters
		LineShape.LineDash = oLineDash				
		oPage.add(LineShape)				
	next i	

	for i = 0 to 4
		Point.x = x
		Point.y = y + (i * M1)
		Size.Width = 8750 ' 1 meters @ 1:80 4/5th of 1250 is 1000/100 = 1 meter 
		Size.Height = 0
		LineShape = oDoc.createInstance("com.sun.star.drawing.LineShape")
		LineShape.Size = Size
		LineShape.Position = Point
		LineShape.LayerID = 6		
		LineShape.LineColor = 0
		LineShape.LineWidth = 10		

		LineShape.LineStyle = com.sun.star.drawing.LineStyle.DASH		
		' Create a new DashLine object
		oLineDash = LineShape.LineDash
		' Set the values of the object
		oLinedash.Style = 1 ' 0 = Dashes 1 & 4 = dots 2 = None huh?
		oLineDash.Dots = 2 
		oLineDash.DotLen = 1
		oLineDash.Dashes = 1
		oLineDash.DashLen = 100
		oLineDash.Distance = 50
		
		' Accept the new parameters
		LineShape.LineDash = oLineDash		
			
		oPage.add(LineShape)	
	next i
end sub

sub a4_setup
	' setup A4 landscape/ working area of 28500 x 19800
	oPage.BorderLeft = 600
	oPage.BorderRight = 600
	oPage.BorderTop = 600
	oPage.BorderBottom = 600
	
	oPage.Width = 29700 
	oPage.Height = 21000
end sub

sub add_border_layer
	' The Layer 5 is named "Borders". For page border lines. If not exist create it.
	oLM = oDoc.getLayerManager()
    lm_count = oLM.getCount()
	'msgbox "Lm_count: " + lm_count
	if lm_count < 6 then
		oLM.insertNewByIndex(5).Name = "Borders"
		oLM.getByIndex(5).IsVisible = True
	end if	
end sub

sub add_grid_layer
	' The Layer 6 is named "Grid". For building base 1 meter grid lines. If not exist create it.
	oLM = oDoc.getLayerManager()
    lm_count = oLM.getCount()
	'msgbox "Lm_count: " + lm_count
	if lm_count < 7 then
		oLM.insertNewByIndex(6).Name = "Grid"
		oLM.getByIndex(6).IsVisible = True
	end if
end sub

sub add_pile_layer
	' The Layer 7 is named "Piles". for displaying the piles If not exist create it.
	oLM = oDoc.getLayerManager()
    lm_count = oLM.getCount()
	'msgbox "Lm_count: " + lm_count
	if lm_count < 8 then
		oLM.insertNewByIndex(7).Name = "Piles"
		oLM.getByIndex(7).IsVisible = True
	end if
end sub

sub border_line
	' Draw a border at 600. Polygon more simple than drawing individual lines.
	Dim PolyPolygonShape As Object
	Dim PolyPolygon As Variant
	Dim Square1(3) As New com.sun.star.awt.Point
	
	PolyPolygonShape = oDoc.createInstance("com.sun.star.drawing.PolyPolygonShape")
	PolyPolygonShape.LayerID = 5		
	PolyPolygonShape.LineColor = 0
	PolyPolygonShape.LineWidth = 10
	PolyPolygonShape.FillTransparence = 100
		 
	oPage.add(PolyPolygonShape) ' Page.add must take place before the coordinates are set
	 
	Square1(0).x = 600
	Square1(1).x = 29100
	Square1(2).x = 29100
	Square1(3).x = 600
	Square1(0).y = 600
	Square1(1).y = 600
	Square1(2).y = 20400
	Square1(3).y = 20400

	PolyPolygonShape.PolyPolygon = Array(Square1())
end sub

sub border_text_field
	' Add a border line at the bottom to make a text field.
	' Text field is to the right, so buttons can be to the left
	' Use rectangle
	Point.x = 10000 '600
	Point.y = 19500
	Size.Width = 18500 + 600 '28500
	Size.Height = 900	

	RectangleShape = oDoc.createInstance("com.sun.star.drawing.RectangleShape")
	RectangleShape.LayerID = 5
	RectangleShape.Size = Size
	RectangleShape.Position = Point
	RectangleShape.LineColor = 0
	RectangleShape.LineWidth = 10
	RectangleShape.CharColor = 0
	
	RectangleShape.FillStyle = com.sun.star.drawing.FillStyle.SOLID
	RectangleShape.LineJoint = com.sun.star.drawing.LineJoint.MITER
	RectangleShape.Name = "Page Label"
	RectangleShape.FillTransparence = 100
	'RectangleShape.FillColor = RGB(255,0,0)	 
	oPage.add(RectangleShape)

	'The text can only be inserted after the drawing object has been added to the drawing page.
	'RectangleShape.String = "A4 Landscape. Scale 1:500"
	' Give it a name so it can be found and changed
	RectangleShape.Name = "Page Label"
	RectangleShape.String =	LABEL
	RectangleShape.CharWeight = com.sun.star.awt.FontWeight.NORMAL 'BOLD
	RectangleShape.CharFontName = "FreeSans" '"Arial"
	RectangleShape.CharHeight = 14	
end sub

sub clear_elements
	' Clear all elements off the drawing, except the Control buttons. 
	' Work backwards, step -1, so removal does not impact the indexing.
	'msgbox "Total element count: " + oPage.getCount
	for i = oPage.getCount - 1 to 0 step -1
		element = oPage.getByIndex(i)
		if element.LayerID = 3 then
			'msgbox "Control element in Layer 3 detected."
		else
			oPage.Remove(element)
		end if
	next i			
end sub

sub clear_pile
	' Clear any previous piles which are on Layer 7
	global_instantiation	
	
	for i = oPage.getCount - 1 to 0 step -1
		element = oPage.getByIndex(i)
		'msgbox element.DBG_Properties
		'msgbox element.control.name + " " + element.control.label + " " + Cstr(i) + " " + Cstr(element.LayerID)
		if element.LayerID = 7 then	
			oPage.Remove(element)
		end if
	next i
end sub
	
sub Pile(x as integer, y as integer)
	' Create a pile of 200mm x 200mm starting at x, y
	' Need to positioning offset if pile size is changed
	' Layer 7 is for piles.

	Dim w as integer
	Dim h as integer
		
	' 100mm = 125 points. Pile is 200mm x 200mm
	Point.x = x
	Point.y = y
	Size.Width = PILE_X_SIZE
	Size.Height = PILE_Y_SIZE

	RectangleShape = oDoc.createInstance("com.sun.star.drawing.RectangleShape")
	RectangleShape.Size = Size
	RectangleShape.Position = Point
	RectangleShape.LineColor = 0
	RectangleShape.LineWidth = 10
	RectangleShape.LayerID = 7
	RectangleShape.FillStyle = com.sun.star.drawing.FillStyle.SOLID

	RectangleShape.FillTransparence = 20
	RectangleShape.FillColor = RGB(0,0,255)	 
	oPage.add(RectangleShape)

end sub

sub debug(optional oObject as object)
	' Utility routine to write an objects properties, methods and supported
	' interfaces to a text file.
	' Calls the function ShellSort()
	' Tested on Ubuntu 20.04.2 and LibreOffice Version: 6.4.7.2
	' Author: Ian Stewart. Date: July 2021
	' Usage: Add debug subroutine and ShellSort function to your Basic code.
	' In your Basic code add a line to call debug. E.g. debug(oDocument)
	dim s(2) as string '3 dimensions to string array
	dim sData as string ' temp string
	dim i as integer	
	dim j as integer
	dim iNum as integer
	dim sLine as String
	dim sMsg as String
	dim sPathFile as String
		
	if IsMissing (oObject) then
		msgbox "No object supplied. Perform testing with 'ThisComponent' object." 
		oObject = ThisComponent
	end if
	
	' Add data to the string array
	s(0) = oObject.Dbg_Properties
	s(1) = oObject.Dbg_Methods
	s(2) = oObject.Dbg_SupportedInterfaces	

	' Properties, Methods and Supported Interfaces
	for j = 0 to 2
		' Split into two parts array. Heading and Data	 
		' Heading is first two lines ends with a colon, ":"
		aTwoPart = split(s(j), ":")
    	l = LBound(aTwoPart)
    	u = UBound(aTwoPart)
		
		' Some objects have no Support Interfaces so there is no data
		if u = 0 then
			sHeading = aTwoPart(0)
			sData = ""
			goto Skip1
		end if
		
		' Make the header string as one line with newline
		sHeading = join(split(aTwoPart(0), chr(10)), " ") + ":" + chr(13)
		
		' Body has random chr(10)'s throughout. Remove:	
		sData = replace(aTwoPart(1), chr(10), " ")
	
		if j = 2 then
			' For Supported Interfaces
			' Remove groups of blank spaces
			for x = 6 to 2 step -1
				sData = replace(sData, space(x), " ")
			next x
			sData = trim(sData)
			sData = replace(sData, " ", chr(13))
			' Use shell sort function to sort the array
			aData = split(sData, Chr(13))
			aSorted_Data = ShellSort(aData)
			sData = join(aSorted_Data, chr(13))			
		end if

		if j < 2 then
			' For properties and methods
			' Create an array based on semi-colon & space, 
			' however in some cases there is semicolon and double space
			aData = split(sData, "; ")
			
	    	l = LBound(aData)
	    	u = UBound(aData)
	
			' Clean up the data one line at a time.
	    	for i = l to u	
				aData(i) = trim(aData(i))
				aData(i) = replace(aData(i), "Sbx", "")
				aData(i) = replace(aData(i), "( ", "(")
				aData(i) = replace(aData(i), " )", ")")
				aData(i) = replace(aData(i), ",", ", ")
				aData(i) = replace(aData(i), "  ", " ")	
			
				' Split the line and make lowercase bracketed field in Methods
				aLinePart = split(aData(i), " (")
				if UBound(aLinePart) = 1 then
					aLinePart(1) = LCase(aLinePart(1))
					aData(i) = join(aLinePart, " (")
				end if
			
				' Split the line and make the first field have a constant length
				aLinePart = split(aData(i), " ")
	    		' Pad first field with spaces.
	    		if len(aLinePart(0)) < 13 then
	    			aLinePart(0) = aLinePart(0) + space(12 - len(aLinePart(0)))
	    		end if
				aData(i) = join(aLinePart, " ")
						
			next i	
			
			' Use shell sort routine to sort the array
			aSorted_Data = ShellSort(aData)
		    l = LBound(aSorted_Data)
		    u = UBound(aSorted_Data)
	    					
			' Convert sorted array into a string with newlines
			sData = join(aSorted_Data(), chr(13))	
	
		end if 
		
		Skip1:
		
		' Output to a file. Needs to be checled to work with MS	
		iNum = Freefile
		' CurDir doesn't work properly. 
		' ie. CurDir + /dgb_info.txt will be HOME folder
		' To do: Test this on MS platform
		
		sPathFile = CurDir + "/dbg_info.txt"		
		
		' Output for Properties, then append for methods and supported interface
		if j = 0 then		
			open sPathFile for output As iNum
		else:
			open sPathFile for append As iNum
		end if

		print #iNum, sHeading + sData + chr(13) + chr(13)		
		close #iNum	
				
	next j
	
	' Open the file with a text editor. E.g. Pluma or Gedit
	if FileExists("/usr/bin/pluma") then
		Shell ("/usr/bin/pluma", 1, sPathFile, false)
	
	elseif FileExists("/usr/bin/gedit") then
		Shell ("/usr/bin/gedit", 1, sPathFile, false)
							
	else:
		' Or post a message
		msgbox "DGB Properties, Methods and Supported Interfaces written to:" + _
				chr(13) + chr(13) + sPathFile
	end if
				
end sub

function ShellSort(mArray)
	' Routine from: https://wiki.openoffice.org/wiki/Sorting_and_searching 
	dim n as integer, h as integer, i as integer, j as integer, t as string, _
			 Ub as integer, LB as integer
	Lb = lBound(mArray)
	Ub = uBound(mArray)	 
	' compute largest increment
	n = Ub - Lb + 1
	h = 1
	if n > 14 then
	        do while h < n
	                h = 3 * h + 1
	        loop
	        h = h \ 3
	        h = h \ 3
	end if
	do while h > 0
	' sort by insertion in increments of h
	        for i = Lb + h to Ub
	                t = mArray(i)
	                for j = i - h to Lb step -h
	                        if strComp(mArray(j), t, 0) < 1 then exit for
	                        mArray(j + h) = mArray(j)
	                next j
	                mArray(j + h) = t
	        next i
	        h = h \ 3
	loop	
	ShellSort = mArray
end function

sub notes
	' Reference information for developers
	
	' Using msgbox...
	s = "The following control will be removed: " + chr(13) 
	if msgbox (s + "Layer ID: " + cstr(element.LayerID) + chr(13) + _
		   "Layer Name: " + element.LayerName + chr(13) + _
		   "Control Name: " + element.control.name + chr(13) + _
		   "Control Label: " + element.control.label, _
		   MB_YESNO + MB_DEFBUTTON1 + MB_ICONQUESTION, _
		   "Remove Control") = IDYES then
		   
		x
	else
		y
	end if
	
	' Constants for msgbox:
	'0, MB_OK - OK button
	'1, MB_OKCANCEL - OK and Cancel button
	'2, MB_ABORTRETRYIGNORE - Abort, Retry, and Ignore buttons
	'3, MB_YESNOCANCEL - Yes, No, and Cancel buttons
	'4, MB_YESNO - Yes and No buttons
	'5, MB_RETRYCANCEL - Retry and Cancel buttons

	'0, MB_DEFBUTTON1 - First button is default value
	'256, MB_DEFBUTTON2 - Second button is default value
	'512, MB_DEFBUTTON3 - Third button is default value				

	'16, MB_ICONSTOP - Stop sign
	'32, MB_ICONQUESTION - Question mark
	'48, MB_ICONEXCLAMATION - Exclamation point
	'64, MB_ICONINFORMATION - Tip icon

	'1, IDOK - Ok
	'2, IDCANCEL - Cancel
	'3, IDABORT - Abort
	'4, IDRETRY - Retry
	'5 - Ignore
	'6, IDYES - Yes
	'7, IDNO - No
	
	
	' Using MRI...
	'If Not Globalscope.BasicLibraries.isLibraryLoaded("MRILib") Then
	'      Globalscope.BasicLibraries.LoadLibrary( "MRILib" )
	'End If
	'Dim oMRI as object
	'oMRI = CreateUnoService( "mytools.Mri" ) 
	
	'          oDoc = ThisComponent
	'oMRI.inspect oDoc
	'          myLibraryForm = oDoc.Drawpage.Forms.getByName("MainForm")
	'          oControl = myLibraryForm.getByName("Author")
	'oMRI.inspect oControl
	'          item = oControl.SelectedItems(0)
	'          book_author = oControl.StringItemList(item)


	' Getting elements
	Dim i
	Dim j
	'Dim element as object	
	'msgbox element.name
	element = oPage.getByIndex(46)
	msgbox element.DBG_Properties
	msgbox CStr(element.LayerID) '3
	msgbox element.LayerName ' controls	
	msgbox element.Title ' Null
	msgbox element.Name ' Null
	msgbox element.control.Name
	msgbox element.control.Label
	
	oControl = element.control
	msgbox oControl.DBG_Properties
	msgbox oControl.Name
	msgbox oControl.Label
	
	' Only elements on Layer3 will have a control object.
	for i = 0 to oPage.getCount - 1
		element = oPage.getByIndex(i)
		'msgbox element.DBG_Properties
		'msgbox element.control.name + " " + element.control.label + " " + _
		'		 Cstr(i) + " " + Cstr(element.LayerID)
		if element.LayerID = 3 then
			' It will have a control object.
			'msgbox element.DBG_Properties
			msgbox "Layer ID ~ Layer Name: " + Cstr(element.LayerID) + " " + _
					element.LayerName
			msgbox element.control.name
		end if		
		'msgbox element.DBG_properties
	'	sMsg = element.DBG_properties
		'msgbox element.ShapeType
	next i		
end sub
