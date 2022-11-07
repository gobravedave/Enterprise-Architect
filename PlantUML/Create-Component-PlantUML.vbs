'
' Script Name: Create Component PlantUML
' Author: David Anderson
' Purpose: Callable routine for dealing with the creation of a component PlantUML script
' Date: 21-Feb-2021
'
sub CreateComponentPlantUML ()
	call LOGInfo("Create Component PlantUML Script activated " & currentDiagram.Stereotype)
	if Left(currentDiagram.Stereotype,2) = "C4" then
		CreateC4PlantUML ()
	else
		LOGWarning("This script does not yet support " & currentDiagram.Type & " diagrams")
	end if
	call LOGInfo ( "Create Component PlantUML Script Complete" )
	call LOGInfo ( "=========================================" )
end sub