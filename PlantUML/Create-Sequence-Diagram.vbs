'
' Script Name: Create Sequence Diagram 
' Author: David Anderson
' Purpose: Sub routine called by the Run PlantUML Script to be used to build a Sequence Diagram  
' Date: 11-Feb-2019
'-----------------------------------------
' Modifcation Log
' 30-Mar-2019:	add logic to support the following
'					- \n for long names
'					- title
'					- dividers (==)
'					- non decalred participants
'					- non space delimitered sequences 
'					- notes left/right and over
'
dim timeline_array (99,7)			'store timeline elements 
dim sequence_array (99,7)			'store interations
dim layout_array (99, 7)			'store coridinates of all sequences and fragments that needs to be positioned
dim t								'timeline array index
dim s								'sequence array index
dim l								'layout array index
dim fragment_level					'fragment level indicator
dim partition_level					'partition level within a fragment
dim left
dim multiline_note					'multi note line indicator is currently being processed and thefore will not hsv
dim note
dim noteName
dim n								'count of note lines
dim autonumber
'dim startNumber
'dim increment

sub CreateSequenceDiagram ()
	call LOGInfo("Create Sequence Diagram script activated")

	dim PlantUML
	dim word

	'check current diagram.. if nothing.. then this script has not been called properly
	
	'the following check is not really required..
	if not theSelectedElement is nothing _
	and theSelectedElement.ObjectType = otElement _
	and theSelectedElement.Type = "Note" then
		'split note..
		'call LOGDebug( "PlantUML")
		dim i
		left=30						'set initial 
		fragment_level=0
		PlantUML = Split(theSelectedElement.Notes,vbcrlf,-1,0)
		for i = 0 to Ubound(PlantUML)
			'call LOGDebug ( "Processing: " & PlantUML(i) )
			if not PlantUML(i) = "" then
				if not Asc(PlantUML(i)) = 39  then
					if multiline_note = True then
						process_note(PlantUML(i))											'process note
					else
						word=split(PlantUML(i))
						select case ucase(word(0))
							case "@STARTUML" 	'call LOGDebug ( "skip: " & PlantUML(i) )		'ignore
							case "AUTONUMBER"	autonumber = True
							case "AUTOACTIVATE"	'call LOGDebug ( "skip: " & PlantUML(i) )		'ignore
							case "TITLE"		create_title(PlantUML(i))
							case "ACTOR"		create_timeline(PlantUML(i))
							case "PARTICIPANT"	create_timeline(PlantUML(i))
							case "BOUNDARY"		create_timeline(PlantUML(i))
							case "CONTROL"		create_timeline(PlantUML(i))
							case "ENTITY"		create_timeline(PlantUML(i))
							case "COLLECTIONS"	create_timeline(PlantUML(i))
							case "DATABASE"		create_timeline(PlantUML(i))
							case "BOX"			create_timeline(PlantUML(i))
							case "END" 			resize_diagramObject(PlantUML(i))			'box or a partition
							case "ACTIVATE" 	'call LOGDebug ( "skip: " & PlantUML(i) )		'ignore
							case "DEACTIVATE" 	'call LOGDebug ( "skip: " & PlantUML(i) )		'ignore
							case "ALT" 			create_fragment(PlantUML(i))				'add fragment
							case "OPT" 			create_fragment(PlantUML(i))				'add fragment
							case "BREAK" 		create_fragment(PlantUML(i))				'add fragment
							case "LOOP" 		create_fragment(PlantUML(i))				'add fragment
							case "CRITICAL" 	create_fragment(PlantUML(i))			 	'add fragment
							case "==" 			create_fragment(PlantUML(i))			 	'add seq fragment as divider						
							case "ELSE" 		add_partition(PlantUML(i)) 					'add partition to fragment
							case "NOTE"			process_note(PlantUML(i))					'process note
							case "@ENDUML" 		'call LOGDebug ( "skip: " & PlantUML(i) )		'ignore
							case else			create_sequence(PlantUML(i))				'replace with a regex expression to make sure sctipt line si indeed a sequence
						end select
					end if
				end if
			end if
		next

		call LOGDebug( "**Timeline Array**" )
		Call PrintArray (timeline_array,0,t-1)
		
		call layout_objects()							'set relative coordinates of seqeunces & fragments
		
		call LOGDebug( "**Layout Array**" )
		Call PrintArray (layout_array,0,l-1)
		
		ReloadDiagram(currentDiagram.DiagramID)
							
		call LOGInfo ( "Create Sequence Diagram Script Complete" )
	else
		call LOGError("problem calling the sub routine")
	end if
end sub

sub create_title(PlantUML)
dim strTitle
dim diagram as EA.Diagram

	call LOGTrace("create_title(" & PlantUML & ")")
	'set the diagram.name using title
		
	strTitle = right(PlantUML, len(PlantUML)-5)
	'remove quotes
	strTitle = replace(strTitle, Chr(34), " ")

	'handle \n
	strTitle = trim(replace(strTitle, "\n", " "))
	
	set diagram = currentDiagram
	diagram.name = strTitle
	diagram.update

	call LOGInfo( "Set Diagram Name to: " & diagram.Name )

end sub

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
dim LOGLEVEL_SAVE	'
'LOGLEVEL_SAVE = LOGLEVEL
'LOGLEVEL=3	

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
							elementStereotype = getStereotype(Plantuml)
	end select
	
	'check if word begins with a quote.. if yes then the name is multi word.. and therefore need to reconstruct name without quotes

	'call LOGDebug( "word(1): " & word(1) & " of: " & ubound(word) )
	if Asc(word(1)) = 34 then		'check for quotes
		if Asc(right(word(1), 1)) = 34 then
			elementName = mid(word(1),2,len(word(1))-2)		
		else
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
		end if
	else
		elementName = word(1)
	end if
	'replace \n with a space
	elementName = replace(elementName, "\n", " ")

	sql ="SELECT Object_ID FROM t_object WHERE t_object.Name =" & chr(34) & elementName & chr(34) & " and t_object.Object_Type = " & chr(34) & elementType & chr(34)
	'call LOGDebug ( "sql: " & sql)

	set elements=Repository.GetElementSet(sql,2)
	'call LOGDebug ("elements returned: " & elements.Count)
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
		diagramObject.BackgroundColor = getColor(PlantUML)
	end if
	diagramObject.Update
	diagramObjects.Refresh
	currentDiagram.Update		

'add to timeline_array
	timeline_array (t,0) = element.ElementID
	timeline_array (t,1) = diagramObject.InstanceID
	timeline_array (t,2) = element.Name
	timeline_array (t,3) = element.Alias
	timeline_array (t,4) = Ucase(word(0)) 'PlantUML participant
	timeline_array (t,5) = diagramObject.left
	timeline_array (t,6) = diagramObject.right
	t=t+1
	call LOGTrace("create_timeline")
	
	'restore logging level	
'	LOGLEVEL=LOGLEVEL_SAVE		

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
dim parsedScript
dim word
dim i
dim j
dim r
dim color
dim script_head
dim script_tail
dim script_return
dim script_source
dim script_target
dim element as EA.Element 
dim connector as EA.Connector
dim LOGLEVEL_SAVE	'
LOGLEVEL_SAVE = LOGLEVEL
'LOGLEVEL=5	

	call LOGTrace( "create sequence(" & script & ")")
	
	'parse for color
	i=instr(script,"[")
	if i > 0 then
		color = getColor(script)
		j=instr(script,"]")
		'remove [#] from script
		'call LOGDebug("[= " & i & " ]=" & j)
		script_head=Mid(script,1,i-1)
		script_tail=Mid(script,j+1)		
		script = script_head & script_tail
		'call LOGDebug( "remove color(" & script & ")")
	end if
	
	'parse to ensure sufficient delimiters to support processing
	if instr(script,"-&gt; &gt;") > 0 then
		script = replace(script, "-&gt; &gt;", " -&gt;&gt; ")
	else
		if instr(script,"-&gt;") > 0 then
			script = replace(script, "-&gt;", " -&gt; ")
		end if
	end if

	if instr(script,"-&lt; &lt;") > 0 then
		script = replace(script, "-&lt; &lt;", " -&lt;&lt; ")
	else
		if instr(script,"-&lt;") > 0 then
			script = replace(script, "-&lt;", " -&lt; ")
		end if
	end if
	
	script = replace(script, ":", " : ")
	script = replace(script, "  ", " ")
	call LOGTrace( "parsedScript: " & script)

	word=split(script)
	
	if ucase(word(0)) = "RETURN" then
		'resolve source and target timelines 
		if s > 0 then
			script_source = timeline_array (timelineIndexById(sequence_array (s-1,2)),2) 	'resolve source using the target element in the prior sequence
			script_target = resolveTarget(sequence_array (s-1,2))							'resolve target using the target element from the prior sequence
			script_return = script_source & " --> " & script_target & " : "					'return in plantuml do not appear to have :
			script = replace(script, word(0), script_return)
			'call LOGDebug( "return script updated to(" & script & ")")
			word=split(script)	
		else
			call LOGWarning(script & ":no where to return")
			exit sub										'skip unresolved return
		end if
	end if
	
	'add to sequence array
	'call LOGDebug( "word count(" & ubound(word)+1 & ")")
	'call LOGDebug( "word1(" & word(1) & ")")
	'check direction of the sequence
	if instr(word(1), "&lt;") = 0 then
		sequence_array (s,0) = timelineElementID(word(0))	'source elementID					
		sequence_array (s,1) = word(1)						'interaction type
		sequence_array (s,2) = timelineElementID(word(2))	'target elementID
	else
		sequence_array (s,2) = timelineElementID(word(0))	'source elementID					
		sequence_array (s,1) = word(1)						'interaction type
		sequence_array (s,0) = timelineElementID(word(2))	'target elementID
	end if
	
	'create connector
	set element = Repository.GetElementByID (sequence_array (s,0))
	set connector = element.Connectors.AddNew("","Sequence")
	
	connector.SupplierID = sequence_array (s,2)
	connector.SequenceNo = (s+1)*10
	if autonumber = True then
		connector.Name = s+1 & ". " & connectorName(script)
	else
		connector.Name = connectorName(script)
	end if
	
	'Check lifecycle
	if instr(script, "**") > 0 then
		connector.Subtype="New"
	end if
	if instr(script, "!!") > 0 then
		connector.Subtype="Delete"
	end if
	
	connector.DiagramID = currentDiagram.DiagramID
	connector.Color = color

	sequence_array (s,3) = connector.Name
	sequence_array (s,4) = synch(word(1))
	sequence_array (s,5) = signature(script)
	sequence_array (s,6) = isReturn(word(1))
	sequence_array (s,7) = connector.ConnectorID
	
	connector.Update
	element.Connectors.Refresh
	
	dim sql
	sql ="UPDATE t_connector SET PData1 = '" & sequence_array (s,4) & "', PData2 = '" & sequence_array (s,5)  & "', PData4 = '" & sequence_array (s,6) & "' WHERE Connector_Id = " & connector.ConnectorID & ";"
	'call LOGDebug ( "SQL: " & sql)
	
	Repository.Execute(sql)
	set connector = Repository.GetConnectorByID(connector.ConnectorID)	
	element.Update
	'call LOGDebug ("+created connector (" & connector.ConnectorID & ")" & vbcrlf & _
	'				vbtab & vbtab & vbtab & " synch: " & connector.MiscData(0) )
	
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
	
	'restore logging level	'
	LOGLEVEL=LOGLEVEL_SAVE		'

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
	i = instr(script, " ") 'postion to 1st char after key word
	if i > 0 then
		fragmentName = mid(script, i+1)
	end if
	
	'remove trialing == of divider
	fragmentName = replace(fragmentName, "==", "")
	fragmentName = trim(fragmentName)
	set elements = currentPackage.Elements
	set element = elements.AddNew( fragmentName, "InteractionFragment" )
	'handle when name is not supplied eg loop

	element.Subtype = fragment_type(mid(script,1, i-1))
	element.Update
	elements.Refresh
	call LOGInfo( "added fragment: " & fragmentName & " (" & element.ElementID & ")" )
	
	'Do not increment fragment level if divider (==) as it will always be top level
	if element.Subtype = 9 then
		layout_array (l,1) = "Divider"			'type of object ie seq, Fragment
	else
		fragment_level=fragment_level+1	
		layout_array (l,1) = "InteractionFragment"	'type of object ie seq, Fragment
	end if
	
	'add fragment to layout array
	layout_array (l,0) = fragment_level			'used to calculate height of object at same level					
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

sub process_note(script)
dim i
dim LOGLEVEL_SAVE	'
LOGLEVEL_SAVE = LOGLEVEL
'LOGLEVEL=4	

	call LOGTrace("process_note (" & script & ")") 

	if ucase(script) = "END NOTE" then
		'add note element
		call add_note()
		multiline_note = False	'switch off multi line note indictor
	else
		if multiline_note = True then
			call contsruct_note(script)
		else	
			i = inStr(script, ":")
			'call LOGDebug( "offset for : = " & i & " of " & len(script))
			'plantuml doesnot recognise note:.. so assume anything after note and before the colon controls layout
			if i > 0 then								'single line note
				noteName = Mid(script, 5, i-5)
				'call LOGDebug( "noteName : = " & noteName)
				call contsruct_note(right(script, len(script)-i))
				call add_note()
			else
				multiline_note = True
				noteName = right(script, len(script)-5)
				note=""
				r = 0
			end if
		end if
	end if
	
	'restore logging level	'
	LOGLEVEL=LOGLEVEL_SAVE		'

end sub

sub contsruct_note(script)
dim notes
dim i

	call LOGTrace("construct_note (" & script & ")") 

	notes = split(script, "\n")
	
	for i = 0 to ubound(notes)
		if note = "" then
			note = notes(0)
		else
			note = note & vbcrlf & notes(i)
		end if
		n=n+1
	next
	'call LOGDebug( "note* (" & note & ":" & n & ")")
	
end sub

sub add_note()
'relies on global variables: note, noteName and n

dim elements as EA.Collection
dim element as EA.Element
dim diagramObjects as EA.Collection
dim diagramObject as EA.DiagramObject
dim diagramObjectName

	call LOGTrace("add_note (" & noteName & ")") 
	'add next element
	set elements = currentPackage.Elements
	set element = elements.AddNew(trim(noteName), "Note" )
	element.Notes = note
	element.Update
	
	'add to layout array
	layout_array (l,0) = n											
	layout_array (l,1) = "Note"					'type of object ie seq, Fragment
	layout_array (l,2) = element.ElementID		'id being element
	l=l+1
	note=""
	n=0

end sub

function timelineElementID(word)
dim i
	call LOGTrace("timelineElementID(" & word & ")")
	if right(word,1) = ":" then
		word = mid(word, 1, len(word)-1)				'remove trailing :
	end if
	timelineElementID=99
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
	'add logic to create timeline if not found in the arrary..
	if timelineElementID = 99 then
		'call LOGDebug( "create a timeline for participant =" & word)
		call create_timeline("participant " & word )
		timelineElementID = timeline_array (t-1,0)
	end if
	
	call LOGTrace("timelineElementID=" & timelineElementID )
	
end function

function getColor(script)
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
	call LOGTrace("getColor(" & script & ")" )
	
	script = replace(script, "[", "[ ")		'add space so as [ to split on # for a sequence
	script = replace(script, "]", " ]")		'add space so as ] to split on # for a sequence

	word=split(script)
	for i = 0 to ubound(word)
		if Asc(word(i)) = 35 then		'check for hash
			'call LOGDebug( "word=" & word(i))
			if ishex(mid(word(i),2,len(word(i))-1)) then
				hexRGB = mid(word(i),2,6)
			else
				hexRGB = mid(ColorHexByName(Ucase(word(i))),2,6)
			end if
			'call LOGDebug( "hexRGB=" & hexRGB)
			exit for
		end if
	next
	if not hexRGB = "" then
		getColor=clng("&h" & mid(hexRGB,5,2) & mid(hexRGB,3,2) & mid(hexRGB,1,2))
		'call LOGDebug( "color decimal=" & getColor)
	end if
	call LOGTrace("getColor=" & getColor )
	
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
	if 	arrow = "--&gt;" or _ 
		arrow = "-->" or _
	 	arrow = "&gt;--" or _ 
		arrow = "<--"then
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
		case "=="			fragment_type = 9		
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
	call LOGTrace("layout_Objects()")

	'LOGDebug ("layout array count l-1=" & l-1 & " Ubound(layout_array)=" & Ubound(layout_array))
	for i = 0 to l-1
		'call calculate heights (reursively) 
		select case layout_array(i,1)
			case "Divider" 						layout_array (i,3) =  dividerHeight (layout_array (i,0), i) 
			case "InteractionFragment" 			layout_array (i,3) =  fragmentHeight (layout_array (i,0), i) 
			case "Partition" 					layout_array (i,3) =  partitionHeight (layout_array (i,0), i) 
			case "Note" 						layout_array (i,3) =  noteHeight (i) 
		end select
	next
	
	'set cordinates of each object
	top = -130

	for i = 0 to l-1	
		'set top as an accumulation of object
		'if layout_array(i,1) = "Title" then 
		'	layout_array (i,4) = 0 
	'		layout_array (i,5) = -140 
	'		'resize based on timeline edges
'			layout_array (i,6) = 30 
'			layout_array (i,7) = 250 
'		else
			layout_array (i,4) = top 
			if layout_array(i,1) = "InteractionFragment"  or _
				layout_array(i,1) = "Divider" then 
				layout_array (i,4) = top + 20
				layout_array (i,5) = layout_array (i,4) - layout_array (i,3)	'bottom
				'set left and right coordinates based on the timelines	
				call setLeftRightCoordinates(layout_array (i,0),i)
			else
				if  layout_array(i,1) = "Note" then
					call setNoteLayout(i)
					top = top - layout_array (i,3)
					layout_array (i,5) = layout_array (i,4) - layout_array (i,3)	'bottom
				else
					top = top - height(layout_array (i,1))
					'self message height
					layout_array (i,5) = layout_array (i,4)
				end if
			end if
'		end if
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
'		if layout_array(i,1) = "Title" then 
'			LOGDebug ("i="& i & ":Processing Title "  & layout_array (i,2)) 
'			'add diagramobject
'			diagramObjectName= "l=" & layout_array (i,6) & ";r=" & layout_array (i,7) & ";t=" & layout_array (i,4) & ";b=" & layout_array (i,3)
'			set diagramObject = currentDiagram.DiagramObjects.AddNew(diagramObjectName, layout_array (i,1))
'			diagramObject.top = layout_array (i,4)
'			diagramObject.bottom = layout_array (i,5)
'			diagramObject.left = layout_array (i,6)
'			diagramObject.right = layout_array (i,7)
'			diagramObject.ElementID = layout_array (i,2)
'			'centre text, fontsize and make bold
'			Call LOGInfo( "created diagramObject (" & diagramObject.ElementID  & ") Top=" & diagramObject.top & " Bottom=" & diagramObject.bottom & " Left=" & diagramObject.left & " Right=" & diagramObject.right )
'			diagramObject.Update
'			diagramObjects.Refresh
'			currentDiagram.Update		
'		end if
		if layout_array(i,1) = "InteractionFragment" or _
			layout_array(i,1) = "Divider" then 
			'LOGDebug ("i="& i & ":Processing InteractionFragment "  & layout_array (i,2)) 
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
			'LOGDebug ("i="& i & ":Processing partition for "  & layout_array (i,2)) 
			dim pcount
			pcount=0
			for j = i-1 to 0 step -1
				'LOGDebug ("j=" & j)
				if layout_array (i,2) = layout_array (j,2) then 'resolve which partition by counting the partitions with the same element id..
					pcount=pcount+1
					'LOGDebug ("increment pcount to " & pcount)
				end if
			next
			set element = Repository.GetElementByID(layout_array (i,2))
			'for each partition in element.Partitions
				'call LOGDebug( "Partitition for " & element.ElementID & " " & "name=" & partition.Name & " object type=" & partition.ObjectType & " size=" & partition.size)	
			'next
			'LOGDebug ("Update partition number " & pcount-1 & " of " & element.Partitions.Count & " to " & layout_array (i,3)) 
			set partition = element.Partitions.GetAt(pcount-1) 
			'LOGDebug ("partition size to be updated from " & partition.Size & " to " & layout_array (i,3)) 
			partition.Size = layout_array (i,3)
			element.Update
		end if
		if layout_array(i,1) = "Note" then
			'LOGDebug ("i="& i & ":Processing Note "  & layout_array (i,2)) 
			'add diagramobject
			diagramObjectName= "l=" & layout_array (i,6) & ";r=" & layout_array (i,7) & ";t=" & layout_array (i,4) & ";b=" & layout_array (i,3)
			set diagramObject = currentDiagram.DiagramObjects.AddNew(diagramObjectName, layout_array (i,1))
			diagramObject.top = layout_array (i,4)
			diagramObject.bottom = layout_array (i,5)
			'set based left/right of
			diagramObject.left = layout_array (i,6)
			diagramObject.right = layout_array (i,7)
			diagramObject.ElementID = layout_array (i,2)
'			'centre text, fontsize and make bold
			Call LOGInfo( "created diagramObject (" & diagramObject.ElementID  & ") Top=" & diagramObject.top & " Bottom=" & diagramObject.bottom & " Left=" & diagramObject.left & " Right=" & diagramObject.right )
			'default color
			set element = Repository.GetElementByID(layout_array (i,2))
			if instr(element.Name,"#") > 0 then
				diagramObject.BackgroundColor = getColor(element.Name)
			end if
			diagramObject.Update
			diagramObjects.Refresh
			currentDiagram.Update		
		end if
	next

	'resize any timeline boxes 
	bottom = 0
	for i = 0 to l-1
		if layout_array (i,5) < bottom then
			bottom = layout_array (i,5)
		end if
	next
	
	call LOGDebug( "*resize timeline boxes to " & bottom)
	for i = 0 to t-1
		'if timeline_array (i,4) = "BOX" then
			call LOGDebug( "*timeline (" & timeline_array(i,0) & ") to be resized")
			for each diagramObject in currentDiagram.DiagramObjects
				if diagramObject.ElementID = timeline_array(i,0) then
					diagramObject.bottom = bottom - 5
					diagramObject.Update
					exit for
				end if
			next
		'end if
	next
	
	currentDiagram.Update
	Repository.SaveDiagram(currentDiagram.DiagramID)
	'Repository.ReloadPackage(currentPackage.PackageID)
	ReloadDiagram(currentDiagram.DiagramID)
	call LOGTrace( "**Layout Array - updated**" )
	Call PrintArray (layout_array,0,l-1)

LOGLEVEL=LOGLEVEL_SAVE

end sub

function height(thing)
	call LOGTrace("height(" & thing & ")")
	select case thing
		case "Note"					height = 20 
		case "Sequence"				height = 35
		case "Sequence2Self"		height = 45
		case "InteractionFragment" 	height = 0		'alt, loop etc
		case "Partition"			height = 0		'else
		case "End"					height = 0		
		case "Divider" 				height = 20		'seq
		case else 					height = 0
	end select
	call LOGTrace("height=" & height)
end function

function dividerHeight(level, start)
dim i
	
	call LOGTrace( "dividerHeight(" & level & ":" & start & ")" )
	dividerHeight=0
	for i = start+1 to l-1 
		if layout_array(i,1) = "Divider" then
			exit for
		end if
		dividerHeight = dividerHeight + height(layout_array (i,1))
	next
	call LOGTrace( "dividerHeight=" & dividerHeight )

end function

function fragmentHeight(level, start)
dim i
	
	call LOGTrace( "fragmentHeight(" & level & ":" & start & ")" )
	fragmentHeight=0
	for i = start to l-1 
		'look for end of fragment for this level
		if layout_array(i,1) = "End" and _
			layout_array(i,0) = level then
			exit for
		end if
		fragmentHeight = fragmentHeight + height(layout_array (i,1))
	next
	call LOGTrace( "fragmentHeight=" & fragmentHeight )

end function

function noteHeight(i)
	
	call LOGTrace( "noteHeight(" & i & ")" )
	noteHeight = height("Note") * layout_array(i,0)
	call LOGTrace( "noteHeight=" & noteHeight )

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

dim LOGLEVEL_SAVE	'
LOGLEVEL_SAVE = LOGLEVEL
'LOGLEVEL=3	

	call LOGTrace("setLeftRightCoordinates(" & level & ":" & start & ")")

	layout_array(start,6) = 400
	layout_array(start,7) = 0

	for i = start to l-1
		'need some way to identify nested fragments and whether the left and right values need adjusting
		if layout_array(i,1) = "Sequence" or _
			layout_array(i,1) = "Sequence2Self" then 
			set connector = Repository.GetConnectorByID(layout_array(i,2))
			
			j=timelineIndexById(connector.ClientID)
			tLeft = timeline_array(j,5)
			'call LOGDebug("client(" & connector.ClientID & ") " & tLeft)
			if tLeft < layout_array(start,6) then
				layout_array(start,6) = tLeft - 25
			end if
			tRight = timeline_array(j,6)
	'		call LOGDebug("client(" & connector.ClientID & ") " & tRight)
			if tRight > layout_array(start,7) then
				layout_array(start,7) = tRight +25
			end if

			if layout_array(i,1) = "Sequence" then
				j=timelineIndexById(connector.SupplierID)	
				tLeft = timeline_array(j,5)
				'call LOGDebug("supplier(" & connector.SupplierID & ") " & tLeft)
				if tLeft < layout_array(start,6) then
					layout_array(start,6) = tLeft -25
				end if
				tRight = timeline_array(j,6)
				'call LOGDebug("supplier(" & connector.SupplierID & ") " & tRight)
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
			'call LOGDebug("Left=" & layout_array(start,6) & ":Right=" & layout_array(start,7))
			
		end if
		'look for end of fragment for this level
		if layout_array(i,1) = "End" and _
			layout_array(i,0) = level then
			exit for
		end if
	next

	call LOGTrace("setLeftRightCoordinates: Left=" & layout_array(start,6) & ":Right=" & layout_array(start,7))
	'restore logging level	'
	LOGLEVEL=LOGLEVEL_SAVE		'

end sub

sub setNoteLayout(i)
dim j
dim word
dim element as EA.Element

	call LOGTrace("setNoteLayout(" & i & ")")
	'default width =155

	'get element
	set element = Repository.GetElementByID(layout_array (i,2))
	'split element.name
	word=split(element.Name)
	if ucase(word(0)) = "LEFT" then
		'get timeline based on name
		j=timelineIndexByName(word(2))	
		'calculate left
		layout_array(i,6) = timeline_array(j,5) - 80
		if layout_array(i,6) < 0 then
			layout_array(i,6) =0
		end if
		'set right based on timeline left
		layout_array(i,7) = timeline_array(j,5) + 35
	else
		if ucase(word(0)) =  "RIGHT" then
			'get timeline based on name
			j=timelineIndexByName(word(2))	
			'calculate left
			layout_array(i,6) = timeline_array(j,6) - 35
			'set right based on timeline left
			layout_array(i,7) = timeline_array(j,6) + 80
		else									
			'over one timeline
			'call LOGDebug("over(" & word(1) & " of " & ubound(word) & ")")
			word(1) = replace(word(1), ",", "")		'remove comma
			j=timelineIndexByName(word(1))	
			'call LOGDebug("t1(" & layout_array(i,4) & "):b1(" & layout_array(i,5) & ")")
			'call LOGDebug("l1(" & layout_array(i,6) & "):r1(" & layout_array(i,7) & ")")
			'call LOGDebug("l2(" & timeline_array(j,5) & "):r2(" & timeline_array(j,6) & ")")
			layout_array(i,6) = timeline_array(j,5) 
			layout_array(i,7) = timeline_array(j,6)
			if ubound(word) > 1 then
				j=timelineIndexByName(word(2))	
				'call LOGDebug("t1(" & layout_array(i,4) & "):b1(" & layout_array(i,5) & ")")
				'call LOGDebug("l1(" & layout_array(i,6) & "):r1(" & layout_array(i,7) & ")")
				'call LOGDebug("l2(" & timeline_array(j,5) & "):r2(" & timeline_array(j,6) & ")")
				if timeline_array(j,5) < layout_array(i,6) then
					layout_array(i,6) = timeline_array(j,5)
				end if
				if timeline_array(j,6) > layout_array(i,7) then
					layout_array(i,7) = timeline_array(j,6)
				end if
			end if
		end if		
	end if
	
end sub

function timelineIndexById(timelineId )
dim i
	call LOGTrace("timelineIndexById(" & timelineId & ")")
	'call LOGDebug("t=" & t)

	for i = 0 to t-1
		if timelineId = timeline_array(i,0) then				'check using element id
			timelineIndexById = i
			exit for
		end if
	next

	call LOGTrace("timelineIndexById=" & timelineIndexById)

end function

function timelineIndexByName(timelineName)
dim i
	call LOGTrace("timelineIndexByName(" & timelineName & ")")
	'call LOGDebug("t=" & t)

	for i = 0 to t-1
		if timelineName = timeline_array(i,2) then				'check using element name
			timelineIndexByName = i
			exit for
		end if
	next

	call LOGTrace("timelineIndexByName=" & timelineIndexByName)

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

function resolveTarget(sourceId)
dim i
	call LOGTrace("resolveTarget(" & sourceId & ")")
	'call LOGDebug("s=" & s)
	resolveTarget=""
	for i = s to 0 step -1
		'call LOGDebug("i=" & i)
		if sourceId = sequence_array(i,2) and sequence_array(i,6) = 0 then		'check using elementid & not a return flow
			resolveTarget = timeline_array (timelineIndexById(sequence_array(i,0)),2) 
			exit for
		end if
	next
	if resolveTarget="" then
		call LOGWarning(sourceId & ":no where to return.. so will treat as a self return")
		resolveTarget = timeline_array(timelineIndexById(sourceId ),2)
	end if
	call LOGTrace("resolveTarget=" & resolveTarget)

end function

function getStereotype(script)
dim start
dim length
	call LOGTrace("getStereotype(" & script & ")")

	if instr(script, "&lt;&lt;") > 0 then
		start = instr(script, "&lt;&lt;")+8
		length = instr(script, "&gt;&gt;") - start
		getStereotype = trim(mid(script, start, length)) 
	end if
	call LOGTrace("getStereotype=" & getStereotype)
	
end function