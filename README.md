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


