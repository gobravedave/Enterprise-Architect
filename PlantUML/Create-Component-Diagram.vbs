'
' Script Name: Create Component Diagram
' Author: David Anderson
' Purpose: Sub routine called by the Run PlantUML Script to build a Component Diagram  
' Date: 29-Mar-2019
'
' 25-Sept-2022:		Support C4 Diagrams
'
sub CreateComponentDiagram ()
	call LOGInfo("Create Component Diagram script activated")
	
	if instr(Ucase(theSelectedElement.Notes),"C4-PLANTUML") > 0 then
		CreateC4Diagram ()
	else
		LOGWarning("This script does not yet support " & currentDiagram.Type & " diagrams")
	end if
	
	call LOGInfo ( "Create Component Diagram Script Complete" )
	call LOGInfo ( "=========================================" )

end sub