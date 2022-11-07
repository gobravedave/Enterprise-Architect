option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Common.Print-Array
!INC Common.Sort-Array
!INC Common.color-picker
!INC EAScriptLib.VBScript-Logging
!INC PlantUML.Create-Activity-PlantUML
!INC PlantUML.Create-Class-PlantUML
!INC PlantUML.Create-Component-PlantUML
!INC PlantUML.Create-Deployment-PlantUML
!INC PlantUML.Create-Sequence-PlantUML
!INC PlantUML.Create-UseCase-PlantUML
!INC PlantUML.Create-C4-PlantUML

'LOGLEVEL=LOGLEVEL_ERROR
'LOGLEVEL=LOGLEVEL_INFO
'LOGLEVEL=LOGLEVEL_WARNING
'LOGLEVEL=LOGLEVEL_DEBUG
LOGLEVEL=LOGLEVEL_TRACE

'
' Script Name: Create PlantUML Script
' Author: David Anderson
' Purpose: Wrapper script to appear in the Diagram Scripting group  
' 		   responsible for directing to the relevant script by diagram type.  
' Date: 11-March-2019
'
' Change Log:
' 18-Sept-2022:		Add C4 Diagram support
'
dim currentDiagram as EA.Diagram
dim currentPackage as EA.Package
Dim selectedObject as EA.DiagramObject
Dim theSelectedElement as EA.Element

dim element_array (99,12)
dim relationship_array (99,4)
dim e				'index for element_array
dim r				'index for relationship_array

sub OnDiagramScript()
	'Show the script output window
	Repository.EnsureOutputVisible "Script"
	call LOGInfo ("------VBScript Generate PlantUML script------" )

	' Get a reference to the current diagram
	set currentDiagram = Repository.GetCurrentDiagram()
	set currentPackage = Repository.GetPackageByID(currentDiagram.PackageID)

	if not currentDiagram is nothing then
		dim selectedObjects as EA.Collection
		set selectedObjects = currentDiagram.SelectedObjects
		if selectedObjects.Count = 1 then
			' One or more diagram objects are selected
			set selectedObject = selectedObjects.GetAt (0)
			set theSelectedElement = Repository.GetElementByID(selectedObject.ElementID)
			if not theSelectedElement is nothing _
				and theSelectedElement.ObjectType = otElement _
				and theSelectedElement.Type = "Note" then
				select case currentDiagram.Type
					case "Activity"		call CreateActivityPlantUML ()
					case "Logical"		call CreateClassPlantUML ()
					case "Component"	call CreateComponentPlantUML ()
					case "Deployment"	call CreateDeploymentPlantUML ()
					case "Sequence"		call CreateSequencePlantUML ()
					case "Use Case"		call CreateUseCasePlantUML ()
					case else			call LOGWarning("This script does not yet support " & currentDiagram.Type & " diagrams")
				end select
				call LOGInfo ( "Script Complete" )
				call LOGInfo ( "===============" )
				Session.Prompt "Done" , promptOK
			else
				Session.Prompt "A note object should be selected" , promptOK
			end if
		else
			if selectedObjects.Count = 0 then
				' Nothing is selected
				Session.Prompt "A note object should be selected" , promptOK
			else
				Session.Prompt "Only one object should be selected" , promptOK				
			end if
		end if
	else
		Session.Prompt "This script requires a diagram to be visible", promptOK
	end if

end sub

OnDiagramScript