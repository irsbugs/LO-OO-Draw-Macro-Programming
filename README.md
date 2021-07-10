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

To access the BASIC script click on: Tools --> Macros --> Edit Macros --> floor_plan.odg --> Standaard --> Module1. An Integrated Developemnt Environment window then provides a means to edit and run the BASIC scripting.














