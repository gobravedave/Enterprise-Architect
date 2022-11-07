'
' Script Name: Create Class Diagram
' Author: David Anderson
' Purpose: Sub routine called by the Run PlantUML Script to build a Class Diagram  
' Date: 29-Mar-2019
'
' 25-Sept-2022:		Support C4 Diagrams
'

dim class_array (99,7)			'store class elements 
dim idxC						'class array index

dim relationship_array (99,7)	'store relationships 
dim idxR						'reltionship array index

sub CreateClassDiagram ()
	call LOGInfo("Create Class Diagram script activated")

	if instr(Ucase(theSelectedElement.Notes),"C4-PLANTUML") > 0 then
		CreateC4Diagram ()
	else
		LOGWarning("This script does not yet support " & currentDiagram.Type & " diagrams")
	end if
	
	call LOGInfo ( "Create Class Diagram Script Complete" )
	call LOGInfo ( "=========================================" )

end sub