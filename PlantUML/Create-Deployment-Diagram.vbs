'
' Script Name: Create Deployment Diagram
' Author: David Anderson
' Purpose: Sub routine called by the Run PlantUML Script to build a Deployment Diagram  
' Date: 29-Mar-2019
'
' 25-Sept-2022:		Support C4 Diagrams
'
sub CreateDeploymentDiagram ()
	call LOGInfo("Create Deployment Diagram script is activated")
	
	if instr(Ucase(theSelectedElement.Notes),"C4-PLANTUML") > 0 then
		CreateC4Diagram ()
	else
		LOGWarning("This script does not yet support " & currentDiagram.Type & " diagrams")
	end if
	
	call LOGInfo ( "Create Deployment Diagram Script Complete" )
	call LOGInfo ( "=========================================" )
	
end sub