'
' Script Name: GeneratePlantUMLScript
' Author: David Anderson
' Purpose: Sub routine called by the Generate PlantUML Script to be used to build a Sequence PlantUML  
' Date: 31-Jan-2019
'
'-----------------------------------------
' Modifcation Log
' 30-Mar-2019:	add logic to support the following
'					- \n for long names
'					- title
'					- dividers
'					- notes
'					- divider
' Diagram Script main function
'
dim timeline_array (99,10)
dim sequence_array (99,8)
dim t				'index for timeline_array
dim s				'index for sequence_array

sub CreateSequencePlantUML()
	call LOGInfo("Create Sequence PlantUML script activated")

	' Get a reference to the current diagram
	dim currentDiagram as EA.Diagram
	set currentDiagram = Repository.GetCurrentDiagram()
	dim generated_script
	
		' Get a reference to any selected objects
	dim selectedObjects as EA.Collection
	set selectedObjects = currentDiagram.SelectedObjects
	dim diagramObject as EA.DiagramObject
	dim element as EA.Element
	dim connector as EA.Connector
	
	' One or more diagram objects are selected
	Dim theSelectedElement as EA.Element
	Dim selectedObject as EA.DiagramObject
	set selectedObject = selectedObjects.GetAt (0)
	set theSelectedElement = Repository.GetElementByID(selectedObject.ElementID)
	'spin through diagram objects and declare participants
	dim line_offset
	t=0
	s=0
	dim partition as EA._Partition
	for each diagramObject in currentDiagram.DiagramObjects
		set Element = Repository.GetElementByID(diagramObject.ElementID)
		if inStr("...Sequence,Actor,Boundary,Component...", element.type) > 0 then
			timeline_array (t,0) = element.ElementID
			timeline_array (t,1) = element.Type
			if instr(element.Name, " ") = 0 then
				timeline_array (t,2) = element.Name
			else
				timeline_array (t,2) = chr(34) & element.Name & chr(34)
			end if
			'replace spaces with \n if length greater than 20 
			if len(timeline_array (t,2)) > 20 then
				timeline_array (t,2) = replace(timeline_array (t,2), " ","\n")
			end if
			if instr(element.Alias, " ") = 0 then
				timeline_array (t,3) = element.Alias	
			else
				timeline_array (t,3) = chr(34) & element.Alias & chr(34)
			end if
			timeline_array (t,4) = element.Stereotype
			timeline_array (t,5) = diagramObject.Left
			timeline_array (t,6) = diagramObject.right
			timeline_array (t,7) = "N"  	'activate switch
			timeline_array (t,8) = lcase(color(element.Type, diagramObject.BackgroundColor))
			timeline_array (t,9) = diagramObject.Left + (diagramObject.right - diagramObject.left)/2  'timeline centre
			t=t+1							
		else
			if element.Type = "InteractionFragment" then
'				call LOGDebug( "*Fragment (" & element.ElementID & ") name=" & element.Name & _
'								" t=" & diagramObject.top & ",b=" & diagramObject.bottom & ",l=" & diagramObject.left & ",r=" & diagramObject.right)
				'add to sequence array
				sequence_array (s,0) = diagramObject.top *-1
				sequence_array (s,1) = 0
				sequence_array (s,2) = 0
				sequence_array (s,3) = element.Name
				sequence_array (s,4) = fragment_type(element.Subtype)
				sequence_array (s,5) = ""
				sequence_array (s,6) = "" 
				s=s+1
				if element.Partitions.Count > 0 then
					line_offset = (diagramObject.top *-1)+ 20
					for each partition in element.Partitions
'						call LOGDebug( "Partitition for " & element.ElementID & " " & "name=" & partition.Name & " object type=" & partition.ObjectType & " operator=" & partition.Operator & " size=" & partition.size & " note=" & partition.Note)	
						sequence_array (s,0) = line_offset
						sequence_array (s,1) = 0
						sequence_array (s,2) = 0
						sequence_array (s,3) = partition.Name
						sequence_array (s,4) = "Else"
						sequence_array (s,5) = ""
						sequence_array (s,6) = "" 
						line_offset = line_offset + partition.Size
						s=s+1
					next
				end if
				'suppress end for dividers
				if not (sequence_array (s-1,4) = "divider") then
					sequence_array (s,0) = diagramObject.bottom *-1
					sequence_array (s,1) = 0
					sequence_array (s,2) = 0
					sequence_array (s,3) = ""
					sequence_array (s,4) = "End"
					sequence_array (s,5) = ""
					sequence_array (s,6) = "" 
					s=s+1
				end if 
			else
				if element.Type = "Note" then 
					if not (element.Name = "PlantUML") then
						'add to sequence array
						sequence_array (s,0) = diagramObject.top *-1
						sequence_array (s,1) = 0
						sequence_array (s,2) = 0
						sequence_array (s,3) = element.Notes
						sequence_array (s,4) = "Note"
						sequence_array (s,5) = diagramObject.BackgroundColor
						sequence_array (s,6) = "" 
						sequence_array (s,7) = diagramObject.left 
						sequence_array (s,8) = diagramObject.right 						
						s=s+1
					else
						call LOGWarning( "PlantUML " & element.type & " element type not added to timeline array")
					end if
				end if
			end if
		end if
	next

	call QuickSort(timeline_array,0,t-1,5)
	call LOGDebug( "Sorted Timeline Array" )
	call PrintArray (timeline_array,0,t-1)
'
	call LOGDebug( "Sequence Array" )
	call PrintArray (sequence_array,0,s-1)
	
	dim box_right
	dim strLine
	box_right=0
	generated_script="@startuml" & vbcrlf & "autoactivate on" & vbcrlf
	generated_script = generated_script & "title " & currentDiagram.Name & vbcrlf
	dim i
	'Output PlantUML participants
	for i = 0 to Ubound(timeline_array)
		if timeline_array (i,0) = "" then
			Exit for
		end if
		if timeline_array (i,1) = "Boundary" then
			strLine = "Box " & timeline_array (i,2) 						
			box_right = timeline_array (i,6)
		Else
			if box_right > 0 then			'check for inline box
				if timeline_array (i,5) > box_right then
					generated_script = generated_script & "End Box" & vbcrlf
					box_right=0
				end if
			end if
			strLine = participant(timeline_array (i,1), timeline_array (i,4)) & " " & timeline_array (i,2) 
			if not timeline_array (i,3) = "" then
				strLine = strLine & " as " & timeline_array (i,3)
			end if
		End if
		if not	timeline_array (i,8) = "" then			'append color if one exists
			strLine = strLine & " " & timeline_array (i,8)
		end if
		generated_script = generated_script & strLine & vbcrlf
	next
	
	if box_right > 0 then				'check for trailing
		generated_script = generated_script & "End Box" & vbcrlf
		box_right=0
	end if

	'spin through diagram links
	dim diagramLink as EA.DiagramLink
'					dim connectorEnd as EA.ConnectorEnd
'					dim connectorTag as EA.ConnectorTag
	
	for each diagramLink in currentDiagram.DiagramLinks
		set connector = Repository.GetConnectorByID(diagramLink.ConnectorID)
		'add to sequence array
'		call LOGDebug( "+Connector (" & connector.ConnectorID & ") " &  connector.Name & _
'					" #:" & connector.SequenceNo & ", sx:" & connector.StartPointX & ", sy:" & connector.StartPointY & ", ex:" & connector.EndPointX & ",ey: " & connector.EndPointY )
		'Session.Output( " type: " & connector.Type )
		'Session.Output( " subtype: " & connector.Subtype )
		'Session.Output( " styleEx: " & connector.StyleEx )
		'Session.Output( " parmam & retval: " & connector.MiscData(1) )
		'Session.Output( " custom property count: " & connector.CustomProperties.Count)
		'Session.Output( " event flags: " & connector.EventFlags)
		'Session.Output( " tags: " & connector.TaggedValues.Count)
		'Session.Output( " states: " & connector.StateFlags)
		'Session.Output( " metatype: " & connector.MetaType)
		
		sequence_array (s,0) = connector.StartPointY *-1					
		sequence_array (s,1) = connector.ClientID
		sequence_array (s,2) = connector.SupplierID
		sequence_array (s,3) = connector.Name
		sequence_array (s,4) = connector.MiscData(0)		'synch or async
		sequence_array (s,5) = connector.MiscData(1)		'return value and parameters
		sequence_array (s,6) = connector.MiscData(3)		'isReturn
		s=s+1
	next
	'sort links from top to bottom
	Call QuickSort(sequence_array,0,s-1,0)
	call LOGDebug( "Sorted Sequence Array" )
	Call PrintArray (sequence_array,0,s-1)

	'Output PlantUML sequences

	for i = 0 to Ubound(sequence_array)
		if sequence_array (i,0) = "" then
			Exit for
		end if
		if sequence_array (i,1) = 0 then		'source/target identifiers are equal 0
			if sequence_array (i,4) = "divider" then
				strline = "== " & sequence_array(i,3) & " ==" 							
			else
				if sequence_array (i,4) = "Note" then
					'use proximity to resolve position
					strline = "note " & resolveNoteLocation(i) & " " & color("Note", sequence_array (i,5)) & vbcrlf & sequence_array(i,3) & vbcrlf & "end note" 											
				else
					strline = sequence_array (i,4) & " " & sequence_array(i,3)
				end if
			end if
			generated_script = generated_script & strLine & vbcrlf
		else
			strline = timeline(sequence_array(i,1)) & arrow(sequence_array (i,4), sequence_array (i,6)) & timeline(sequence_array(i,2))  
			if not sequence_array(i,3) = "" then
				strline = strline & ": " & sequence_array(i,3) & signature(sequence_array (i,5))
			end if
			generated_script = generated_script & strLine & vbcrlf

			'activate source if synchronous and not already active
'			if sequence_array (i,4) = "Synchronous" then
'				if activate_timeline(sequence_array(i,1)) = "Y" then
'					strline = "activate " & timeline(sequence_array(i,1))
'					generated_script = generated_script & strLine & vbcrlf
'				end if
'			end if

			'activate target if asynchronous and not already active
'			if activate_timeline(sequence_array(i,2)) = "Y" then
'				strline = "activate " & timeline(sequence_array(i,2))
'				generated_script = generated_script & strLine & vbcrlf
'			end if
			
			'deactivate source
'			if sequence_array (i,6) = 1 then			'isreturn
'				deactivate_timeline(sequence_array(i,1))
'				strline = "deactivate " & timeline(sequence_array(i,1))
'				generated_script = generated_script & strLine & vbcrlf
'			end if
		end if
	next
	
	'deactivate any active timelines
'	For i = 0 to t-1
'		if timeline_array (i,0) = "" then
'			Exit for
'		end if
'		if timeline_array (i,7) = "Y" then
'			timeline_array (i,7) = "N"
'			if timeline_array(i,3) = "" then 
'				strline = "deactivate " & timeline_array(i,2)
'			else
'				strline = "deactivate " & timeline_array(i,3)
'			end if
'			generated_script = generated_script & strLine & vbcrlf							
'		end if
'	Next
	'check for InteractionFragments
	
	generated_script=generated_script & "@enduml"
	theSelectedElement.Notes = generated_script
	theSelectedElement.Update

end sub

function participant(strType, strStereotype)
	If strType = "Actor" then
		participant = strType
	else
		select case Ucase(strStereotype)
			case "DATABASE" 	participant = "Database"
			case "BOUNDARY" 	participant = "Boundary"
			case "CONTROL" 		participant = "Control"
			case "ENTITY" 		participant = "Entity"
			case "COLLECTIONS" 	participant = "Collections"
			case else			participant = "Participant" 
		end select
	end if
end function

function timeline(elementid)
Dim i
	Call LOGTrace( "timeline(" & elementid & ")" )

	For i = 0 to Ubound(timeline_array)
		if timeline_array (i,0) = "" then
			Exit for
		end if
		if timeline_array (i,0) = elementid then
			if timeline_array (i,3) = "" then
				timeline = timeline_array (i,2)		'return name
			else
				timeline = timeline_array (i,3)		'return alias
			end if
			'check if spaces.. if enclosed in quotes
			Exit for
		end if
	Next
	Call LOGTrace( "timeline=" & timeline )

end function

'function activate_timeline(elementid)
'	Dim i
'	For i = 0 to Ubound(timeline_array)
'		if timeline_array (i,0) = "" then
'			Exit for
'		end if
'		if timeline_array (i,0) = elementid then
'			if timeline_array (i,7) = "N" then
'				timeline_array (i,7) = "Y"
'				activate_timeline = "Y"		'activate
'			else
'				activate_timeline = "N"		'already activated
'			end if
'			Exit for
'		end if
'	Next
'end function

'maynot be required if autoactivate is set
'function deactivate_timeline(elementid)
'Dim i
'	For i = 0 to Ubound(timeline_array)
'		if timeline_array (i,0) = "" then
'			Exit for
'		end if
'		if timeline_array (i,0) = elementid then
'			if timeline_array (i,7) = "Y" then
'				timeline_array (i,7) = "N"
'				deactivate_timeline = "Y"		'deactivated
'			else
'				deactivate_timeline = "N"		'not activated
'			end if
'			Exit for
'		end if
'	Next
'end function

function arrow(misc0, misc3)
	Call LOGTrace( "arrow(" & misc0 & ":" & misc3 & ")" )
	
	if misc0 = "Asynchronous" then
		arrow = " ->> "
	else
		if isnull(misc3) then	
			arrow = " -> "				'synchronous
		else
			if misc3 = 0 then	
				arrow = " -> "				'synchronous
			else
				if misc3 = 1 then			' is return?
					arrow = " --> "
				end if
			end if
		end if
	end if
	
	Call LOGTrace( "arrow=" & arrow )
	
end function

function signature(misc2)
	'parse miscdata2 for params and retrun value
	call LOGTrace( "signature(" & misc2 & ")")

	dim i 
	dim j
	dim l
	dim retval
	dim param
	
	if misc2="" then
		signature = " ()"
		exit function
	end if
	
	i = inStr(misc2, "retval=void")
	if i = 0 then
		i = inStr(misc2, "retval=")
		if i > 0 then
			j = inStr(i, misc2, chr(59))
			l=j-(i+7)
			if j>0 then
				retval=mid(misc2,i+7, l)
			end if
		end if
	end if

	i = instr(misc2,"paramsDlg=") 
	if i > 0 then
		j = instr(i, misc2,chr(59))
		l = j-(i+10)
		param=mid(misc2,i+10,l)
	end if
	
	if param = "" then
		param="()"
	else
		param = "(" & param & ")" 
	end if
	
	if not retval = "" then
		retval=":" & retval
	end if

	signature = param & retval
	call LOGTrace( "Signature=" & signature )
	
end function

function color(elementType, BackgroundColor)
'receives decimal version of rbg
'resolve default value if value passed is -1
'resolve ColorNameByHex

dim hexvalue
dim hexRGB
	call LOGTrace("color(" & elementType & ":" & BackgroundColor & ")")
	if BackgroundColor = -1 then
		select case elementType
			case "Sequence" 	color = "#lightblue"
			case "Component" 	color = "#lightpink"
			case "Boundary" 	color = "#lightgrey"
			case "Note" 		color = "#lightyellow"			
			case else 			color=""
			end select
	else
		'Session.Output( "background color=" & BackgroundColor)
		hexvalue = hex(BackgroundColor)
		while len(hexvalue) < 6
			hexvalue = "0" & hexvalue
		wend		
		hexRGB = "#" & mid(hexvalue,5,2) & mid(hexvalue,3,2) & mid(hexvalue,1,2)
		color = ColorNameByHex (hexRGB)
		if color="" then
			color = hexRGB
		end if
'		call LOGDebug( "hexColor=" & color)
	end if 
	call LOGTrace("color=" & color)
	
end function

function fragment_type(ftype)
	select case ftype
		case 0 		fragment_type = "alt"
		case 1 		fragment_type = "opt"
		case 2 		fragment_type = "break"
		case 3 		fragment_type = "par"
		case 4 		fragment_type = "loop"
		case 5 		fragment_type = "critical"
		case 9		fragment_type = "divider"
		case else 	fragment_type=""
	end select
end function

function resolveNoteLocation(sequence_index)
'this funtion does not cater for a Note spanning 2 timelines
dim i		'index for spinning thru time line arrary
dim j    	'index for the selected timeline entry
dim rightSide_gap
dim leftSide_gap
dim gap
dim side
rightSide_gap = 999999
leftSide_gap = 999999

	call LOGTrace("resolveNoteLocation(" & sequence_index & ")")
'	call LOGDebug( "Nbr of timelines=" & t)
	for i = 0 to t-1
'		call LOGDebug( "processing timeline(" & i & ")=" & timeline_array(i, 2) & ":l=" & timeline_array(i, 5) & " r=" & timeline_array(i, 6) & " c=" & timeline_array(i, 9))
		if timeline_array (i,0) = "" then
			Exit for
		end if
		if not (timeline_array (i,1) = "Boundary") then
'			call LOGDebug( "Note l=" & sequence_array(sequence_index,7) & ":r=" & sequence_array(sequence_index,8))
			if sequence_array(sequence_index,7) => timeline_array(i, 9) then
				'note is on the right side of current timeline
				gap = sequence_array(sequence_index,7) - timeline_array(i, 9) 	'calculate leftside_gap
'				call LOGDebug( "Note is on the right side of " & timeline_array(i, 2) & " with a gap of " & gap)
				if gap < leftside_gap then
					side = "right"
					leftside_gap = gap
					j=i
'					call LOGDebug( "Setting timeline(" & timeline_array(i, 2) & ") as being the closest")
				end if	
			end if
			if sequence_array(sequence_index,8) <= timeline_array(i, 9) then
				'note is on the left side of current timeline
				gap = timeline_array(i, 9) - sequence_array(sequence_index,8) 	'calculate rightside_gap
'				call LOGDebug( "Note is on the left side of " & timeline_array(i, 2) & " with a gap of " & gap)
				if gap < rightside_gap then
					side = "left"
					rightside_gap = gap
					j=i
'					call LOGDebug( "Setting timeline(" & timeline_array(i, 2) & ") as being the closest")
				end if
			end if
		else
'			call LOGDebug( "skipping boundary")
		end if
	next
	resolveNoteLocation = side & " of " & timeline(timeline_array(j, 0))
	call LOGTrace("resolveNoteLocation=" & resolveNoteLocation)
end function