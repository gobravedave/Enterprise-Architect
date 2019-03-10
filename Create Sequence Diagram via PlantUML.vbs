option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Common.Print-Array
!INC Common.color-picker
!INC EAScriptLib.VBScript-Logging
'LOGLEVEL=0		'ERROR
LOGLEVEL=1		'INFO
'LOGLEVEL=2		'WARNING
'LOGLEVEL=3		'DEBUG
'LOGLEVEL=4		'TRACE
'
' Script Name: GenerateSequenceDiagramViaPlantUML
' Author: David Anderson
' Purpose: Import PlantUML Script into an EA diagram  
' Date: 11-Feb-2019
'
' Diagram Script main function
'
dim currentDiagram as EA.Diagram
dim currentPackage as EA.Package
dim timeline_array (99,7)			'store timeline elements 
dim sequence_array (99,7)			'store interations
dim layout_array (99, 7)			'store coridinates of all sequences and fragments that needs to be positioned

dim left
dim t								'timeline array index
dim s								'sequence array index
dim f								'fragment array index
dim l								'layout array index

dim fragment_level					'fragment level indicator
dim partition_level					'partition level within a fragment

sub OnDiagramScript()
	'Show the script output window
	Repository.EnsureOutputVisible "Script"
	call LOGInfo ("------VBScript Create Sequence Diagram via PlantUML------" )

	' Get a reference to the current diagram
	set currentDiagram = Repository.GetCurrentDiagram()
	set currentPackage = Repository.GetPackageByID(currentDiagram.PackageID)
	dim PlantUML
	dim word
	
	if not currentDiagram is nothing then
	'check..'
		if currentDiagram.Type = "Sequence" then
			' Get a reference to any selected objects
			dim selectedObjects as EA.Collection
			set selectedObjects = currentDiagram.SelectedObjects
			'dim diagramObject as EA.DiagramObject
			'dim element as EA.Element
			
			if selectedObjects.Count = 1 then
				' One or more diagram objects are selected
				Dim theSelectedElement as EA.Element
				Dim selectedObject as EA.DiagramObject
				set selectedObject = selectedObjects.GetAt (0)
				set theSelectedElement = Repository.GetElementByID(selectedObject.ElementID)
				if not theSelectedElement is nothing and theSelectedElement.ObjectType = otElement and theSelectedElement.Type = "Note" then
					'split note..
					call LOGDebug( "PlantUML")
					dim i
					left=30						'set initial 
					fragment_level=0
					PlantUML = Split(theSelectedElement.Notes,vbcrlf,-1,0)
					for i = 0 to Ubound(PlantUML)
						call LOGDebug ( "Processing: " & PlantUML(i) )
						if not PlantUML(i) = "" then
							if not Asc(PlantUML(i)) = 39  then
								word=split(PlantUML(i))
								select case ucase(word(0))
									case "ACTOR"		create_timeline(PlantUML(i))
									case "PARTICIPANT"	create_timeline(PlantUML(i))
									case "BOUNDARY"		create_timeline(PlantUML(i))
									case "CONTROL"		create_timeline(PlantUML(i))
									case "ENTITY"		create_timeline(PlantUML(i))
									case "COLLECTIONS"	create_timeline(PlantUML(i))
									case "DATABASE"		create_timeline(PlantUML(i))
									case "BOX"			create_timeline(PlantUML(i))
									case "END" 			resize_diagramObject(PlantUML(i))			'box or a partition
									case "ACTIVATE" 	'Session.Output( "skip: " & PlantUML(i) )	'ignore
									case "DEACTIVATE" 	'Session.Output( "skip: " & PlantUML(i) )	'ignore
									case "ALT" 			create_fragment(PlantUML(i))				'add fragment
									case "OPT" 			create_fragment(PlantUML(i))				'add fragment
									case "BREAK" 		create_fragment(PlantUML(i))				'add fragment
									case "LOOP" 		create_fragment(PlantUML(i))				'add fragment
									case "CRITICAL" 	create_fragment(PlantUML(i))			 	'add fragment
									case "ELSE" 		add_partition(PlantUML(i) ) 				'add partition to fragment
									case else			create_sequence(PlantUML(i))				'replace with a regex expression to make sure sctipt line si indeed a sequence
								end select
							end if
						end if
					next

					call LOGTrace( "**Timeline Array**" )
					Call PrintArray (timeline_array,0,t-1)
					
					call layout_objects							'set relative coordinates of seqeunces & fragments
					
					call LOGTrace( "**Layout Array**" )
					Call PrintArray (layout_array,0,l-1)
					
					ReloadDiagram(currentDiagram.DiagramID)
										
					call LOGInfo ( "Script Complete" )
					call LOGInfo ( "===============" )
					Session.Prompt "Done" , promptOK
				else
					Session.Prompt "A note object with the valid PlantUML script should be selected" , promptOK
				end if
			else	
				if selectedObjects.Count = 0 then
					' Nothing is selected
					Session.Prompt "A note object with the valid PlantUML script should be selected" , promptOK
				else
					Session.Prompt "Only one object should be selected" , promptOK					
				end if
			end if		
		else
			Session.Prompt "This script does not yet support " & currentDiagram.Type & " diagrams" , promptOK
		end if
	else
		Session.Prompt "This script requires a diagram to be visible", promptOK
	end if

end sub

OnDiagramScript

sub create_timeline(PlantUML)
dim i
dim word
dim sql
dim elementName
dim elementType
dim elementStereotype
dim elements as EA.Collection
dim element as EA.Element
dim diagramObjects as EA.Collection
dim diagramObject as EA.DiagramObject
dim width
dim diagramObjectName
	call LOGTrace("create_timeline(" & PlantUML & ")")
	word=split(PlantUML)
	select case Ucase(word(0))
		case "BOX"			elementType = "Boundary"
							width=10
							elementStereotype = ""
		case "ACTOR"		elementType = "Actor"
							width=90
							elementStereotype = ""
		case "DATABASE" 	elementType = "Component"
							width=90
							elementStereotype = "database"
		case "BOUNDARY" 	elementType = "Sequence"
							width=50
							elementStereotype = "boundary"
		case "CONTROL" 		elementType = "Sequence"
							width=50
							elementStereotype = "control"
		case "ENTITY" 		elementType = "Sequence"
							width=50
							elementStereotype = "entity"
		case "COLLECTIONS" 	elementType = "Sequence"
							width=90
							elementStereotype = "collections"
		case else			elementType = "Sequence" 
							width=90
							elementStereotype = ""
	end select
	
	'check if word begins with a quote.. if yes then the name is multi word.. and therefore need to reconstruct name without quotes

	if Asc(word(1)) = 34 then		'check for quotes
		elementName = mid(word(1),2)
		'loop until find the word with the enclosing chr(34)
		for i = 2 to Ubound(word)
			if Asc(right(word(i), 1)) = 34 then
				dim x 
				x = int(len(word(i)))-1
				elementName = elementName & " " & mid(word(i),1, x)
				exit for
			else
				elementName = elementName & " " & word(i)
			end if
		next
	else
		elementName = word(1)
	end if
	
	sql ="SELECT Object_ID FROM t_object WHERE t_object.Name =" & chr(34) & elementName & chr(34) & " and t_object.Object_Type = " & chr(34) & elementType & chr(34)
	call LOGDebug ( "sql: " & sql)

	set elements=Repository.GetElementSet(sql,2)
	call LOGDebug ("elements returned: " & elements.Count)
	if elements.Count = 0 then
		'add element
		set elements = currentPackage.Elements
		set element = elements.AddNew( elementName, elementType )
		element.Stereotype = elementStereotype
		if instr(Ucase(PlantUML)," AS ") > 0 then
			element.Alias = elementAlias(PlantUML)
		end if
		if element.Type = "Boundary" then
			dim borderStyle
			set borderStyle = element.Properties("BorderStyle")
			borderStyle.value = "Solid"
		end if
		element.Update
		currentPackage.elements.Refresh
		call LOGInfo( "Added Element: " & element.Name & " (ID=" & element.ElementID & ") in package " & currentPackage.Name )
	else
		set element = Repository.GetElementByID(elements.GetAt(0).ElementID)
	end if

	
	if elementType="Boundary" then
		diagramObjectName= "l=" & left & ";r=" & left+width & ";t=28;b=-278"
	else
		diagramObjectName= "l=" & left & ";r=" & left+width & ";t=50;b=-250"
	end if 
	left=left+width+45
	set DiagramObjects = currentDiagram.DiagramObjects
	set diagramObject = currentDiagram.DiagramObjects.AddNew(diagramObjectName, elementType)
	diagramObject.ElementID = element.ElementID
	'default color
	if instr(PlantUML,"#") > 0 then
		diagramObject.BackgroundColor = color(PlantUML)
	end if
	diagramObject.Update
	diagramObjects.Refresh
	currentDiagram.Update		

'add to tineline_array
	timeline_array (t,0) = element.ElementID
	timeline_array (t,1) = diagramObject.InstanceID
	timeline_array (t,2) = element.Name
	timeline_array (t,3) = element.Alias
	timeline_array (t,4) = Ucase(word(0)) 'PlantUML participant
	timeline_array (t,5) = diagramObject.left
	timeline_array (t,6) = diagramObject.right
	t=t+1
	call LOGTrace("create_timeline")
end sub

sub resize_diagramObject(script)
dim i
dim diagramObject as EA.DiagramObject
	call LOGTrace ("resize_diagramObject(" & script & ")")
	if ucase(script) = "END BOX" then
		'find box in the timeline_array
		for i = t to 0 step -1
			if timeline_array (i,4) = "BOX" then
				for each diagramObject in currentDiagram.DiagramObjects
					if diagramObject.ElementID = timeline_array(i,0) then
						diagramObject.right = left	
						diagramObject.Update
						left=left+15
						exit for
					end if
				next
			end if
		next
	else
		'add fragment end to layout array
		layout_array (l,0) = fragment_level	'used to calculate height of object at same level					
		layout_array (l,1) = "End"			'type of object 
		'layout_array (l,2) = ""			'id of the connector
		l=l+1
		fragment_level = fragment_level-1
	end if
	call LOGTrace ("resize_diagramObject")

end sub

sub create_sequence(script)
dim word
dim i
dim element as EA.Element 
dim connector as EA.Connector

	call LOGTrace( "create sequence(" & script & ")")
	word=split(script)
	
	'add to sequence array
	'Session.Output( "word count(" & ubound(word)+1 & ")")
	sequence_array (s,0) = timelineElementID(word(0))	'source elementID					
	sequence_array (s,1) = word(1)						'interaction type
	sequence_array (s,2) = timelineElementID(word(2))	'target elementID

	'create connector
	set element = Repository.GetElementByID (sequence_array (s,0))
	set connector = element.Connectors.AddNew("","Sequence")
	connector.SupplierID = sequence_array (s,2)
	connector.SequenceNo = (s+1)*10
	connector.Name = connectorName(script)
	connector.DiagramID = currentDiagram.DiagramID

	sequence_array (s,3) = connector.Name
	sequence_array (s,4) = synch(word(1))
	sequence_array (s,5) = signature(script)
	sequence_array (s,6) = isReturn(word(1))
	sequence_array (s,7) = connector.ConnectorID
	
	connector.Update
	element.Connectors.Refresh
	
	dim sql
	sql ="UPDATE t_connector SET PData1 = '" & sequence_array (s,4) & "', PData2 = '" & sequence_array (s,5)  & "', PData4 = '" & sequence_array (s,6) & "' WHERE Connector_Id = " & connector.ConnectorID & ";"
	call LOGDebug ( "SQL: " & sql)
	
	Repository.Execute(sql)
	set connector = Repository.GetConnectorByID(connector.ConnectorID)	
	element.Update
	call LOGDebug ("+created connector (" & connector.ConnectorID & ")" & vbcrlf & _
					vbtab & vbtab & vbtab & " synch: " & connector.MiscData(0) )
	'Session.Output( " signature: " & connector.MiscData(1))
	'Session.Output( " isReturn: " & connector.MiscData(3))
	'Session.Output( " startpoint x: " & connector.StartPointX )
	'Session.Output( " startpoint y: " & connector.StartPointY )
	'Session.Output( " endpoint x: " & connector.EndPointX )
	'Session.Output( " endpoint y: " & connector.EndPointY )
	'Session.Output( " name: " & connector.Name )
	'Session.Output( " type: " & connector.Type )
	'Session.Output( " objecttype: " & connector.ObjectType)
	'Session.Output( " styleEx: " & connector.StyleEx )
	
	'add sequence to layout array
	layout_array (l,0) = fragment_level			'used to calculate height of object at same level
	if connector.ClientID = connector.SupplierID then
		layout_array (l,1) = "Sequence2Self"
	else
		layout_array (l,1) = "Sequence"			'type of object 
	end if
	layout_array (l,2) = connector.ConnectorID	'id of the connector
	l=l+1
	s=s+1
	
end sub

sub create_fragment(script)

dim elements as EA.Collection
dim element as EA.Element
dim diagramObjects as EA.Collection
dim diagramObject as EA.DiagramObject
dim diagramObjectName
dim i
dim fragmentName

	'create element
	i = instr(script, " ")
	if i > 0 then
		fragmentName = mid(script, i+1)
	end if
	set elements = currentPackage.Elements
	set element = elements.AddNew( fragmentName, "InteractionFragment" )
	element.Subtype = fragment_type(mid(script,1, i-1))
	element.Update
	elements.Refresh
	call LOGInfo( "added fragment: " & fragmentName & " (" & element.ElementID & ")" )

	fragment_level=fragment_level+1	

	'add fragment to layout array
	layout_array (l,0) = fragment_level			'used to calculate height of object at same level					
	layout_array (l,1) = "InteractionFragment"	'type of object ie seq, Fragment
	layout_array (l,2) = element.ElementID		'id being element
	l=l+1
	
end sub

sub add_partition(script)
dim element as EA.Element
dim elementId
dim diagramObject as EA.DiagramObject
dim partitions as EA.Collection
dim partition as EA._Partition
dim partitionName
dim i
	i = instr(script, " ")
	if i > 0 then
		partitionName = mid(script, i+1)
	end if

	'get fragment element using the id stored in the layout array
	for i = l to 0 step-1
		if layout_array (i,1) = "InteractionFragment" then
			elementId = layout_array (i,2)
			exit for
		end if
	next
	'
	set element = Repository.GetElementByID(elementId)
	set partitions = element.Partitions
	set partition = partitions.AddNew(partitionName,"Partition")
	partition.Size=40		'default
	element.Update
	currentPackage.elements.Refresh
	call LOGInfo( "add partition: " & partitionName & " to Fragment (" & element.ElementID & ")" )

	'add partition to layout array
	layout_array (l,0) = fragment_level			'used to calculate height of object at same level					
	layout_array (l,1) = "Partition"			'type of object ie seq, Fragment
	layout_array (l,2) = element.ElementID		'id being element
	l=l+1
	
end sub

function timelineElementID(word)
dim i
	call LOGTrace("timelineElementID(" & word & ")")
	if right(word,1) = ":" then
		word = mid(word, 1, len(word)-1)				'remove trailing :
	end if
	
	for i = 0 to ubound(timeline_array)
		if word = timeline_array(i,3) then				'check using alias
			timelineElementID = timeline_array (i,0)
			exit for
		else
			if word = timeline_array(i,2) then			'check using name
				timelineElementID = timeline_array (i,0)
				exit for
			end if
		end if
	next
	call LOGTrace("timelineElementID=" & timelineElementID )
	
end function

function color(script)
'color is delimited by #
'the color value can be either
'	- a Hex version of RGB 
'	- or a standard color name (refer to color-picker script)
'if a color name is provided, call color-picker to return hex value of RGB
'if not.. look up get the hex vales by color name
'return a decimal value of the RGB
'
dim word
dim hexRGB
dim i
	call LOGTrace("color(" & script & ")" )

	word=split(script)
	for i = 0 to ubound(word)
		if Asc(word(i)) = 35 then		'check for hash
			call LOGDebug( "word=" & word(i))
			if ishex(mid(word(i),2,len(word(i))-1)) then
				hexRGB = mid(word(i),2,6)
			else
				hexRGB = mid(ColorHexByName(Ucase(word(i))),2,6)
			end if
			call LOGDebug( "hexRGB=" & hexRGB)
			exit for
		end if
	next
	if not hexRGB = "" then
		color=clng("&h" & mid(hexRGB,5,2) & mid(hexRGB,3,2) & mid(hexRGB,1,2))
		call LOGDebug( "color decimal=" & color)
	end if
	call LOGTrace("color=" & color )
	
end function

function ishex(word)
	ishex= isnumeric("&h" & word)
end function

function elementAlias(script)
dim i
dim word
	elementAlias=""
	
		'check if word begins with a quote.. if yes then the name is multi word.. and therefore need to reconstruct name without quotes

	word=split(script)
	' find  as 
	for i = 0 to ubound(word)
		if ucase(word(i)) = "AS" then
			exit for
		end if
	next
	
	if Asc(word(i+1)) = 34 then		'check for quotes
		elementAlias = mid(word(i+1),2)
		'loop until find the word with the enclosing chr(34)
		for i = i+2 to Ubound(word)
			if Asc(right(word(i), 1)) = 34 then
				dim x 
				x = int(len(word(i)))-1
				elementAlias = elementAlias & " " & mid(word(i),1, x)
				exit for
			else
				elementAlias = elementAlias & " " & word(i)
			end if
		next
	else
		elementAlias = word(i+1)
	end if
end function

function connectorName(script)	
'start with :
'end with (
dim x
dim y

	x = inStr(script, ":")
	y = inStr(script, "(")
	if x > 0 then
		if y > x then
			connectorName = trim(mid(script, x+1, y-x-1))
		else
			connectorName = trim(mid(script, x+1))		
		end if
	end if
end function

function synch(arrow)
	call LOGTrace("synch(" & arrow & ")")
	if 	arrow = "-&gt;&gt;" or _		
		arrow = "->>" then
		synch = "Asynchronous"
	else
		synch = "Synchronous"
	end if
	call LOGTrace("synch=" & synch)
end function

function signature(script)
dim i
dim j
	call LOGTrace("signature(" & script & ")")

	signature=""
	
	i = inStr(script, ":")					'find first :
	j= inStr(i+1,script, ":")				'find second : denoting retrun value
	if j > 0 then
		signature = "retval=" & trim(mid(script,j+1))
	else
		signature = "retval=void"
	end if

	i = inStr(script, "()")					'indicates there are no params
	if i = 0 then
		i = inStr(script, "(")
		if i > 0 then						'add parms
			j = inStr(script, ")")
			signature = signature & ";params=;paramsDlg=" & trim(mid(script, i+1, j-i-1)) & ";"
		end if
	end if
	
	call LOGTrace("signature(" & signature & ")")
	
end function

function isReturn(arrow)
	call LOGTrace("isReturn(" & arrow & ")")
	if 	arrow = "--&gt;&gt;" or _ 
		arrow = "-->>"then
		isReturn = 1
	else
		isReturn = 0
	end if
	call LOGTrace("isReturn=" & isReturn)
	
end function

function fragment_type(script)

	select case script
		case "alt" 			fragment_type = 0
		case "opt" 			fragment_type = 1
		case "break" 		fragment_type = 2
		case "par" 			fragment_type = 3
		case "loop" 		fragment_type = 4
		case "critical"		fragment_type = 5
		case else 			fragment_type = 0
	end select
end function

sub layout_objects()

dim i
dim j
dim connector as EA.Connector
dim element as EA.Element
dim diagramObjectName
dim diagramObjects as EA.Collection
dim diagramObject as EA.DiagramObject
dim diagramLink as EA.DiagramLink

dim partition as EA._Partition
dim top
dim bottom
dim LOGLEVEL_SAVE
LOGLEVEL_SAVE = LOGLEVEL
'LOGLEVEL=3		'DEBUG

	'LOGDebug ("layout array count l-1=" & l-1 & " Ubound(layout_array)=" & Ubound(layout_array))
	for i = 0 to l-1
		'call calculate fragment heights (reursively) 
		if layout_array(i,1) = "InteractionFragment" then 
			layout_array (i,3) =  fragmentHeight (layout_array (i,0), i) 
		else
			if layout_array(i,1) = "Partition" then 
				layout_array (i,3) = partitionHeight (layout_array (i,0), i) 
			else
				'layout_array (i,3) = layout_array (i,3) + height(layout_array (i,1))
				'j = sequenceIndex(layout_array (i,2))
				'if sequence_array(j,0) = sequence_array(j,2) then
				'	layout_array (i,3) = layout_array (i,3) + 10		'adjust hieght
				'end if
			end if
		end if
	next
	
	'set cordinates of each object
	top = -130

	for i = 0 to l-1	
		'set top as an accumulation of object
		layout_array (i,4) = top 
		if layout_array(i,1) = "InteractionFragment" then 
			layout_array (i,4) = top + 20
			layout_array (i,5) = layout_array (i,4) - layout_array (i,3)	'bottom
			'set left and right coordinates based on the timelines	
			call setLeftRightCoordinates(layout_array (i,0),i)
		else
			layout_array (i,4) = top 
			top = top - height(layout_array (i,1))
			'self message height
			layout_array (i,5) = layout_array (i,4)
		end if
	next

	set diagramObjects = currentDiagram.DiagramObjects
	for i = 0 to l-1
		if layout_array(i,1) = "Sequence" then 
			'get connector & update connector Y coordinates
			'set connector = Repository.GetConnectorByID (layout_array (i,2))			
			'set element = Repository.GetElementByID(connector.ClientID)
			'
			'element.Connectors.Refresh
			'set connector = Repository.GetConnectorByID (layout_array (i,2))			
			'call LOGDebug( "*move sequence(" & layout_array (i,2) & ") " & connector.Name & " from " & connector.StartPointX & ":" & connector.StartPointY & " to " & layout_array (i,4))
			'connector.StartPointX = 1		
			'connector.StartPointY = layout_array (i,4)		
			'connector.EndPointX = 2
			'connector.EndPointY = layout_array (i,5)
			'connector.Update
			'element.Connectors.Refresh
			'element.Update
			'currentPackage.elements.Refresh
			'create diagramlink
			'set diagramLink = currentDiagram.DiagramLinks.AddNew("","")
			'diagramLink.ConnectorID=connector.ConnectorID	
			'Call LOGInfo( "created diagramLink (" & diagramLink.ConnectorID  & ")")
			'Session.Output( " geometry: " & diagramLink.Geometry )
			'diagramLink.Update
			'currentDiagram.DiagramLinks.Refresh
			'currentDiagram.Update
			'Call LOGInfo( "created diagramLink (" & diagramLink.ConnectorID  & ")")
			'call LOGInfo( "connector (" & layout_array (i,2) & ") " & connector.Name & " startX:Y=" & connector.StartPointX & ":" & connector.StartPointY)
		end if
		if layout_array(i,1) = "InteractionFragment" then 
			LOGDebug ("i="& i & ":Processing InteractionFragment "  & layout_array (i,2)) 
			'add diagramobject
			diagramObjectName= "l=" & layout_array (i,6) & ";r=" & layout_array (i,7) & ";t=" & layout_array (i,4) & ";b=" & layout_array (i,3)
			set diagramObject = currentDiagram.DiagramObjects.AddNew(diagramObjectName, layout_array (i,1))
			diagramObject.top = layout_array (i,4)
			diagramObject.bottom = layout_array (i,5)
			diagramObject.left = layout_array (i,6)
			diagramObject.right = layout_array (i,7)
			diagramObject.ElementID = layout_array (i,2)
			Call LOGInfo( "created diagramObject (" & diagramObject.ElementID  & ") Top=" & diagramObject.top & " Bottom=" & diagramObject.bottom & " Left=" & diagramObject.left & " Right=" & diagramObject.right )
			diagramObject.Update
			diagramObjects.Refresh
			currentDiagram.Update		
		end if
		if layout_array(i,1) = "Partition" then
			'update partition size
			LOGDebug ("i="& i & ":Processing partition for "  & layout_array (i,2)) 
			dim pcount
			pcount=0
			for j = i-1 to 0 step -1
				LOGDebug ("j=" & j)
				if layout_array (i,2) = layout_array (j,2) then 'resolve which partition by counting the partitions with the same element id..
					pcount=pcount+1
					LOGDebug ("increment pcount to " & pcount)
				end if
			next
			set element = Repository.GetElementByID(layout_array (i,2))
			for each partition in element.Partitions
				call LOGDebug( "Partitition for " & element.ElementID & " " & "name=" & partition.Name & " object type=" & partition.ObjectType & " size=" & partition.size)	
			next
			LOGDebug ("Update partition number " & pcount-1 & " of " & element.Partitions.Count & " to " & layout_array (i,3)) 
			set partition = element.Partitions.GetAt(pcount-1) 
			LOGDebug ("partition size to be updated from " & partition.Size & " to " & layout_array (i,3)) 
			partition.Size = layout_array (i,3)
			element.Update
			LOGLEVEL=LOGLEVEL_SAVE
		end if
	next

	'resize any timline boxes 
	bottom = 0
	for i = 0 to l-1
		if layout_array (i,5) < bottom then
			bottom = layout_array (i,5)
		end if
	next
	
'	call LOGDebug( "*resize timeline boxes to " & bottom)
	for i = 0 to t-1
		if timeline_array (i,4) = "BOX" then
'			call LOGDebug( "*Box (" & timeline_array(i,0) & ") to be resized")
			for each diagramObject in currentDiagram.DiagramObjects
				if diagramObject.ElementID = timeline_array(i,0) then
					diagramObject.bottom = bottom - 5
					diagramObject.Update
					exit for
				end if
			next
		end if
	next
	
	currentDiagram.Update
	Repository.SaveDiagram(currentDiagram.DiagramID)
	'Repository.ReloadPackage(currentPackage.PackageID)
	ReloadDiagram(currentDiagram.DiagramID)
'LOGLEVEL=3
	call LOGTrace( "**Layout Array - updated**" )
	Call PrintArray (layout_array,0,l-1)

'LOGLEVEL=2

end sub

function height(thing)
	call LOGTrace("height(" & thing & ")")
	select case thing
		case "Sequence"				height = 35
		case "Sequence2Self"		height = 45
		case "InteractionFragment" 	height = 0		'alt, loop etc
		case "Partition"			height = 0		'else
		case "End"					height = 0		
		case else 					height = 0
	end select
	call LOGTrace("height=" & height)
end function

function fragmentHeight(level, start)
dim i
	
	call LOGTrace( "fragmentHeight(" & level & ":" & start & ")" )
	fragmentHeight=0
'	for i = start to Ubound(layout_array) 
	for i = start to l-1 
'		if layout_array(i,0) = "" then
'			exit for
'		end if
		'look for end of fragment for this level
		if layout_array(i,1) = "End" and _
			layout_array(i,0) = level then
			exit for
		end if
		fragmentHeight = fragmentHeight + height(layout_array (i,1))
	next
	call LOGTrace( "fragmentHeight=" & fragmentHeight )

end function

function partitionHeight(level, start)
dim i
dim LOGLEVEL_SAVE
	
	LOGLEVEL_SAVE = LOGLEVEL 
'	LOGLEVEL=4			'activate debugging for the sub

	call LOGTrace( "partitionHeight(" & level & ":" & start & ")" )
	partitionHeight=height(layout_array (start,1))
	for i = start+1 to Ubound(layout_array)
		'end of partition is not specifically declared..  
		if layout_array(i,0) = level then
			if layout_array(i,1) = "InteractionFragment" or _
				layout_array(i,1) = "Partition" or _
				layout_array(i,1) = "End" then
				exit for
			end if
		end if
		'is dependant upon nuumber of sequences
		partitionHeight = partitionHeight + height(layout_array (i,1))
		'end of partition is identified when the level indicator is less than what was passed to it 
	next
	call LOGTrace( "partitionHeight=" & partitionHeight )

	'restore logging level	'
	LOGLEVEL=LOGLEVEL_SAVE
	
end function

sub setLeftRightCoordinates(level,start)
'loop thru list of timelines in scope of this level and retrun the lowest value

dim i
dim j
dim connector as EA.Connector
dim tLeft
dim tRight
'LOGLEVEL=3			'

	call LOGTrace("setLeftRightCoordinates(" & level & ":" & start & ")")

	layout_array(start,6) = 400
	layout_array(start,7) = 0

	for i = start to l-1
		'need some way to identify nested fragments and whether the left and right values need adjusting
		if layout_array(i,1) = "Sequence" or _
			layout_array(i,1) = "Sequence2Self" then 
			set connector = Repository.GetConnectorByID(layout_array(i,2))
			
			j=timelineIndex(connector.ClientID)
			tLeft = timeline_array(j,5)
			call LOGDebug("client(" & connector.ClientID & ") " & tLeft)
			if tLeft < layout_array(start,6) then
				layout_array(start,6) = tLeft - 25
			end if
			tRight = timeline_array(j,6)
	'		call LOGDebug("client(" & connector.ClientID & ") " & tRight)
			if tRight > layout_array(start,7) then
				layout_array(start,7) = tRight +25
			end if

			if layout_array(i,1) = "Sequence" then
				j=timelineIndex(connector.SupplierID)	
				tLeft = timeline_array(j,5)
				call LOGDebug("supplier(" & connector.SupplierID & ") " & tLeft)
				if tLeft < layout_array(start,6) then
					layout_array(start,6) = tLeft -25
				end if
				tRight = timeline_array(j,6)
		'		call LOGDebug("supplier(" & connector.SupplierID & ") " & tRight)
				if tRight > layout_array(start,7) then
					layout_array(start,7) = tRight + 25
				end if
			end if
			
			if layout_array(start,6) < 5 then			'left coordinate cannot be less than zero
				layout_array(start,6)=5
			end if 

			'check if margins need to adjusted becuase the fragment is nested
			if level > 1 then
				for j = start to 0 step -1
					if layout_array(j,0) = level -1 and _
						layout_array(j,1) = "InteractionFragment" then		'scan layout arrary to find previous level fragmant
						if layout_array(start,4) = layout_array(j,4) then	'if top value are the same subtract 5
							layout_array(start,4) = layout_array(start,4) - 5
						end if
						if layout_array(start,5) = layout_array(j,5) then	'if bottom values are the same add 5
							layout_array(start,5) = layout_array(start,5) + 5
						end if						
						if layout_array(start,6) = layout_array(j,6) then	'if left value are the same add 10
							layout_array(start,6) = layout_array(start,6) + 10
						end if
						if layout_array(start,7) = layout_array(j,7) then	'if right values are the same subtract 10
							layout_array(start,7) = layout_array(start,7) - 10
						end if						
						exit for
					end if
				next
			end if
			call LOGDebug("Left=" & layout_array(start,6) & ":Right=" & layout_array(start,7))
			
		end if
		'look for end of fragment for this level
		if layout_array(i,1) = "End" and _
			layout_array(i,0) = level then
			exit for
		end if
	next

	call LOGTrace("setLeftRightCoordinates: Left=" & layout_array(start,6) & ":Right=" & layout_array(start,7))
'LOGLEVEL=2			'

end sub

function timelineIndex(timelineId)
dim i
	call LOGTrace("timelineIndex(" & timelineId & ")")
	'call LOGDebug("t=" & t)

	for i = 0 to t-1
		if timelineId = timeline_array(i,0) then				'check using element id
			timelineIndex = i
			exit for
		end if
	next

	call LOGTrace("timelineIndex=" & timelineIndex)

end function

function sequenceIndex(sequenceId)
dim i
	call LOGTrace("sequenceIndex(" & sequenceId & ")")
	'call LOGDebug("s=" & s)

	for i = 0 to s-1
		if sequenceId = sequence_array(i,7) then				'check using element id
			sequenceIndex = i
			exit for
		end if
	next

	call LOGTrace("sequenceIndex=" & sequenceIndex)

end function