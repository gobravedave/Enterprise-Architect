option explicit

!INC Local Scripts.EAConstants-VBScript
!INC EAScriptLib.VBScript-Logging

'LOGLEVEL=0		'ERROR
LOGLEVEL=1		'INFO
'LOGLEVEL=2		'WARNING
'LOGLEVEL=3		'DEBUG
'LOGLEVEL=4		'TRACE
'
' Script Name: ExportPlantUMLscript
' Author: David Anderson
' Purpose: Create a PUML file from the Selected Note Element
' Date: 25-Mar-2019
'
dim currentDiagram as EA.Diagram
dim currentPackage as EA.Package

sub OnDiagramScript()
	'Show the script output window
	Repository.EnsureOutputVisible "Script"
	call LOGInfo ("------VBScript Export PlantUML Script------" )

	' Get a reference to the current diagram
	set currentDiagram = Repository.GetCurrentDiagram()
	set currentPackage = Repository.GetPackageByID(currentDiagram.PackageID)
	dim selectedObjects as EA.Collection
	set selectedObjects = currentDiagram.SelectedObjects
	dim diagramObject as EA.DiagramObject
	dim element as EA.Element
	if selectedObjects.Count = 1 then
		Dim theSelectedElement as EA.Element
		Dim selectedObject as EA.DiagramObject
		set selectedObject = selectedObjects.GetAt (0)
		set theSelectedElement = Repository.GetElementByID(selectedObject.ElementID)
		'check a note element is selected
		if not theSelectedElement is nothing _ 
			and theSelectedElement.ObjectType = otElement _ 
			and theSelectedElement.Type = "Note" then
			dim PlantUMLfn
			dim project
			set project = Repository.GetProjectInterface()
			dim OFN_OVERWRITEPROMPT
			OFN_OVERWRITEPROMPT = &H2
			PlantUMLfn = project.GetFileNameDialog (currentDiagram.Name & ".puml", "PlantUML Files|*.pu;*.puml", 1, OFN_OVERWRITEPROMPT ,"", 1) 
			If PlantUMLfn = "" Then 
				call LOGInfo("File not selected" )
				stop
			Else
				call LOGInfo ("PlantUML Script file selected: " & PlantUMLfn )
				Dim fileSystemObject
				dim outputFile
				call LOGDebug("""" & PlantUMLfn & """")
				' Define Global File IO Objects
				set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
				dim strRow
				If fileSystemObject.FileExists(PlantUMLfn) Then
				  Set outputFile = fileSystemObject.OpenTextFile(PlantUMLfn, 2, True)
				Else
				  Set outputFile = fileSystemObject.CreateTextFile(PlantUMLfn, True)
				End If
				strRow = Split(theSelectedElement.Notes,vbcrlf,-1,0)
				dim i
				for i = 0 to ubound(strRow)
					outputFile.writeline (strRow(i))
					call LOGDebug(strRow(i))
				next
				outputFile.Close
			End If
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
						
	call LOGInfo ( "Script Complete" )
	call LOGInfo ( "===============" )

end sub

OnDiagramScript