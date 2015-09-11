import sys, win32com.client
import copy
import re
import os

# Visio constants
visCharacterColor  = 1
visCharacterFont = 0
visSectionCharacter = 3
visCharacterSize = 7
visCharacterDblUnderline = 8
visSectionFirstComponent = 10

visSectionObject =  1 
visRowPrintProperties =  25 

visPrintPropertiesPageOrientation =  16 
visRowPage =  10 
visPageWidth =  0 
visPageHeight =  1 

# Visio must be open and I used a "Basic Diagram" template
visio = win32com.client.Dispatch("Visio.Application")

FlowchartTemplateName = "Basic Flowchart.vst"
docFlowTemplate = visio.Documents.Add(FlowchartTemplateName)

pg = docFlowTemplate.Pages.Item(1)

# Change page from landscape to portrait but this works sometimes
visio.Application.ActiveWindow.Page.PageSheet.CellsSRC(visSectionObject, visRowPrintProperties, visPrintPropertiesPageOrientation).FormulaForceU = "1"
visio.Application.ActiveWindow.Page.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageWidth).FormulaU = "8.5 in"
visio.Application.ActiveWindow.Page.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageHeight).FormulaU = "11 in"


# Place a Visio shape on the Visio document
def dropShape (shapeType, posX, posY, theText):

    print "Shape type = %s" % shapeType
    print "X = %i" % posX
    print "Y = %i" % posY

    vsoShape = pg.Drop(shapeType, posX, posY)

    setDefaultShapeValues(vsoShape)
    vsoShape.Text = theText

    return vsoShape   # Returns the shape that was created

# Draw connector from bottom of one shape to another shape with autoroute
def connectShapes(shape1, shape2, theText):
    conn = visio.Application.ConnectorToolDataObject
    shpConn = pg.Drop(conn, 0, 0)

    shpConn.CellsU("BeginX").GlueTo(shape1.CellsU("PinX"))          
    shpConn.CellsU("EndX").GlueTo(shape2.CellsU("PinX"))

    setDefaultShapeValues(shpConn)
    shpConn.Text = theText

# Specify which part of a shape to draw a connector from 
#     one shape to another shape.
def connectShapes2(shape1, shape2, glueBegin, glueEnd, theText):

    conn = visio.Application.ConnectorToolDataObject
    shpConn = pg.Drop(conn, 0, 0)

    shpConn.CellsU("BeginX").GlueTo(shape1.CellsU(glueBegin))
    shpConn.CellsU("EndX").GlueTo(shape2.CellsU(glueEnd))

    setDefaultShapeValues(shpConn)
    shpConn.Text = theText

# Set the default color, font and ect for a shape
def setDefaultShapeValues(vsoShape):

    vsoShape.Cells("LineColor").FormulaU = 0
    vsoShape.Cells("LineWeight").FormulaU = "2.0 pt"
    vsoShape.FillStyle  = "None"
    vsoShape.Cells("Char.size").FormulaU = "12 pt"

    vsoShape.CellsSRC(visSectionCharacter, 0, visCharacterDblUnderline).FormulaU = False
    vsoShape.CellsSRC(visSectionCharacter, 0, visCharacterColor).FormulaU = "THEMEGUARD(RGB(0,0,0))"
    vsoShape.CellsSRC(visSectionCharacter, 0, visCharacterFont).FormulaU = 100

    return vsoShape

# Get the stencil object
def getStencilName(): # Name of Visio stencil containing shapes

    FlowchartStencilName = "BASFLO_U.VSSX" # Basic Flow Chart
    docFlowStencil = ""

    for doc in visio.Documents:
        print "Doc name = %s" % doc
        if doc.Name == FlowchartStencilName or doc.Name == "BASFLO_M.VSSX" :

            docFlowStencil = doc

    print "docFlowStencil = %s" % docFlowStencil  # Print installed stencils
    return docFlowStencil

def main(x):

    MasterProcessName = "Process"
    MasterDecisionName = "Decision"
    MasterStartEnd = "Start/End"

    docFlowStencil = getStencilName()

    # Get masters for Process and Decision:
    mstProcess = docFlowStencil.Masters.ItemU(MasterProcessName)
    mstDecision =  docFlowStencil.Masters.ItemU(MasterDecisionName)
    mstStartEnd = docFlowStencil.Masters.ItemU(MasterStartEnd)

    x = 3
    y = 10

    start1  = dropShape(mstStartEnd, x, y, "Start") # Start/stop shape with rounded corners
    process2 = dropShape(mstProcess, x, y - 1.5, "Shape 2") # Process shape with rectangle bpx
    process3 = dropShape(mstProcess, x + 2, y - 3.0, "Shape 3") 
    decide1 = dropShape(mstDecision, x , y - 3.0, "Decide") # Process shape with rectangle bpx
    process4 = dropShape(mstProcess, x, y - 4.5, "Shape 4") 
    end1 = dropShape(mstStartEnd, x, y - 6.0, "End")
  
    # Add connectors to the shapes
    connectShapes(start1, process2, "")
    connectShapes(process2, decide1, "")
    connectShapes(decide1, process4, "YES")
    connectShapes2(decide1, process3, "Connections.X2", "PinX", "NO" )
    connectShapes(process4, end1, "")
    connectShapes(process3, end1, "")

main(1)
