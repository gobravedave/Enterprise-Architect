'
' Script Name: Create C4 Diagram 
' Author: David Anderson
' Purpose: Sub routine called by the Run PlantUML Script to be used to build a C4 Diagram  
' Date: 25-Sep-2022
'-----------------------------------------'

dim stereotype_array (99,7)		'store stereotypes 
dim idxS						'stereotype array index
dim idxL						'index for layout_array 

sub CreateC4Diagram ()
	call LOGInfo("Create C4 Diagram script activated")

	dim PlantUML
	dim word
	dim i
	
	PlantUML = Split(theSelectedElement.Notes,vbcrlf,-1,0)
	for i = 0 to Ubound(PlantUML)
		call LOGDebug ( "Processing #" & i & " :" & PlantUML(i) )
		if not PlantUML(i) = "" then
			PlantUML(i) = Trim(PlantUML(i))
			'C4 commands are
			PlantUML(i) = replace(PlantUML(i), "(", " (")
			if not Asc(PlantUML(i)) = 39  then
				word=split(PlantUML(i))
				select case ucase(word(0))
					case "TITLE"							call create_title(PlantUML(i))
					'ignore
					case "!INCLUDE" 						call LOGWarning ( "skipping: " & PlantUML(i) )		
					case "@STARTUML" 						call LOGWarning ( "skipping: " & PlantUML(i) )		
					case "@ENDUML" 							call LOGWarning ( "skipping: " & PlantUML(i) )
					case "LAYOUT_TOP_DOWN"					call LOGWarning ( "skipping: " & PlantUML(i) )
					case "LAYOUT_AS_SKETCH"					call LOGWarning ( "skipping: " & PlantUML(i) )
					case "LAYOUT_WITH_LEGEND"				call LOGWarning ( "skipping: " & PlantUML(i) )
					case "SHOW_LEGEND"						call LOGWarning ( "skipping: " & PlantUML(i) )
					' not yet supported
					case "RELINDEX"							call LOGWarning ( "not yet supported: " & PlantUML(i) )
					case "INCREMENT"						call LOGWarning ( "not yet supported: " & PlantUML(i) )
					case "SETINDEX"							call LOGWarning ( "not yet supported: " & PlantUML(i) )
					case "LASTINDEX"						call LOGWarning ( "not yet supported: " & PlantUML(i) )					
					case "NODE"								call LOGWarning ( "not yet supported: " & PlantUML(i) )
					case "NODE_L"							call LOGWarning ( "not yet supported: " & PlantUML(i) )
					case "NODE_R"							call LOGWarning ( "not yet supported: " & PlantUML(i) )
					case "DEPLOYMENT_NODE"					call LOGWarning ( "not yet supported: " & PlantUML(i) )	
					case "ADDPROPERTY"						call LOGWarning ( "not yet supported: " & PlantUML(i) )	
					case "WITHOUTPROPERTYHEADER" 			call LOGWarning ( "not yet supported: " & PlantUML(i) )	
					case "ADDNODETAG" 						call LOGWarning ( "not yet supported: " & PlantUML(i) )	
					'class
					case "PERSON"							call create_class(PlantUML(i))
					case "SYSTEM"							call create_class(PlantUML(i))
					case "SYSTEM_EXT"						call create_class(PlantUML(i))
					case "SYSTEMDB"							call create_class(PlantUML(i))
					case "SYSTEMDB_EXT"						call create_class(PlantUML(i))
					case "SYSTEMQUEUE"						call create_class(PlantUML(i))
					case "SYSTEMQUEUE_EXT"					call create_class(PlantUML(i))
					case "CONTAINER"						call create_class(PlantUML(i))
					case "CONTAINER_EXT"					call create_class(PlantUML(i))
					case "CONTAINERDB"						call create_class(PlantUML(i))
					case "CONTAINERDB_EXT"					call create_class(PlantUML(i))
					case "CONTAINERQUEUE"					call create_class(PlantUML(i))
					case "CONTAINERQUEUE_EXT"				call create_class(PlantUML(i))
					case "COMPONENT"						call create_class(PlantUML(i))
					case "COMPONENT_EXT"					call create_class(PlantUML(i))
					case "COMPONENTDB"						call create_class(PlantUML(i))
					case "COMPONENTDB_EXT"					call create_class(PlantUML(i))
					case "COMPONENTQUEUE"					call create_class(PlantUML(i))
					case "COMPONENTQUEUE_EXT"				call create_class(PlantUML(i))
					'boundary
					case "BOUNDARY"							call create_class(PlantUML(i))
					case "SYSTEM_BOUNDARY"					call create_class(PlantUML(i))
					case "CONTAINER_BOUNDARY"				call create_class(PlantUML(i))
					case "ENTERPRISE_BOUNDARY"				call create_class(PlantUML(i))
					case "}"								call end_of(PlantUML(i))
					'sterotype
					case "ADDPERSONTAG"						call LOGWarning ( "skipping: " & PlantUML(i) )
					case "ADDBOUNDARYTAG"					call LOGWarning ( "skipping: " & PlantUML(i) )
					case "ADDELEMENTTAG"					call LOGWarning ( "skipping: " & PlantUML(i) )				
					case "ADDCOMPONENTTAG"					call LOGWarning ( "skipping: " & PlantUML(i) )					
					case "ADDRELTAG"						call LOGWarning ( "skipping: " & PlantUML(i) )	
					case "ADDEXTERNALPERSONTAG" 			call LOGWarning ( "skipping: " & PlantUML(i) )	
					case "ADDEXTERNALSYSTEMTAG" 			call LOGWarning ( "skipping: " & PlantUML(i) )	
					case "ADDEXTERNALCONTAINERTAG" 			call LOGWarning ( "skipping: " & PlantUML(i) )
					case "ADDEXTERNALCOMPONENTTAG" 			call LOGWarning ( "skipping: " & PlantUML(i) )	
					case "UPDATECONTAINERBOUNDARYSTYLE" 	call LOGWarning ( "skipping: " & PlantUML(i) )	
					case "UPDATESYSTEMBOUNDARYSTYLE" 		call LOGWarning ( "skipping: " & PlantUML(i) )
					case "UPDATEENTERPRISEBOUNDARYSTYLE" 	call LOGWarning ( "skipping: " & PlantUML(i) )
					case "UPDATEBOUNDARYSTYLE" 				call LOGWarning ( "skipping: " & PlantUML(i) )
					case "UPDATERELSTYLE" 					call LOGWarning ( "skipping: " & PlantUML(i) )
					'relationship
					case else								call create_relationship(PlantUML(i))				
				end select
			end if
		end if
	next

	call LOGDebug( "**Class Array**" )
	Call PrintArray (class_array,0,idxC-1)

	'apply stereotype settings
	call LOGDebug( "**Stereotype Array**" )
	Call PrintArray (stereotype_array,0,idxS-1)

	call LOGDebug( "**Relationship Array**" )
	Call PrintArray (relationship_array,0,idxR-1)
	
	'layout objects based
	call layout_objects()							'set relative coordinates based nesting of ojects
	'resize boundary elements

	call LOGDebug( "**Layout Array**" )
	Call PrintArray (layout_array,0,idxL-1)
	
	call build_diagram()							'
	
	'LayoutDiagram(currentDiagram.DiagramID)
	ReloadDiagram(currentDiagram.DiagramID)
	call LOGInfo ( "Create C4 Diagram Script Complete" )
	
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

sub create_class(script)
	dim i
	dim word
	dim sql
	dim className
	dim classAlias
	dim classObjectType
	dim classTech
	dim classStereotype
	dim classC4Type
	dim classNotes
	dim elements as EA.Collection
	dim element as EA.Element
	dim elementTag as EA.TaggedValue
	dim LOGLEVEL_SAVE	'
	'LOGLEVEL_SAVE = LOGLEVEL
	'LOGLEVEL=3	

	call LOGTrace("create_class(" & script & ")")
	word=split(script)
	
	if instr(word(0),"Boundary") > 0 then
		classObjectType = "Boundary"
		classStereotype = replace(word(0),"_"," ")
	else
		classObjectType = "Class"
		if instr(word(0),"Container") > 0 then
			'classObjectType = "Container"
			classStereotype = "Container"
		else
			if instr(word(0),"Component") > 0 then
				'classObjectType = "Component"
			else
				if instr(word(0),"System") > 0 then
					'classObjectType = "System"
					classStereotype = "System"
				else
				end if
			end if
		end if
	end if
	
	if instr(word(0),"Db") > 0 then
		classC4Type="Storage"
	else
		if instr(word(0),"Queue") > 0 then
			classC4Type="Queue"
		else
			if instr(word(0),"System") > 0 then
				classC4Type="Software"
			else
				if instr(word(0),"Container") > 0 then
					classC4Type="Application"
				end if
			end if
		end if
	end if
	
	'extract macro which is delimitered by ()
	dim strMacro
	strMacro = trim(Mid(script, instr(script,"(")+1, instr(script,")")-instr(script,"(")-1))
	call LOGDebug( "strMacro: " & strMacro )
	
	dim strArgs
	strArgs=split(strMacro,",")
	for i = 0 to Ubound(strArgs)
		call LOGDebug( "strArgs(" & i & " of " & ubound(strArgs) & ") : " & strArgs(i) )
	next
	classAlias = trim(strArgs(0))
	className = trim(strArgs(1))
	'replace \n with a space
	className = replace(className, "\n", " ")
	className = replace(className, """", "")

	if Ubound(strArgs) > 1 then
		classTech = trim(strArgs(2))
		classTech = replace(classTech, """", "")
	else
		classTech = ""
	end if
	'assumes last arg contains the description
	if Ubound(strArgs) > 2 then
		classNotes = trim(strArgs(Ubound(strArgs)))
		classNotes = replace(classNotes, """", "")
	else
		classNotes = ""
	end if

	sql ="SELECT Object_ID FROM t_object WHERE t_object.Name =" & chr(34) & className & chr(34) & " and t_object.Object_Type = " & chr(34) & classObjectType & chr(34)
	call LOGDebug ( "sql: " & sql)

	set elements=Repository.GetElementSet(sql,2)
	call LOGDebug ("elements returned: " & elements.Count)
	if elements.Count = 0 then
		'add element
		set elements = currentPackage.Elements
		set element = elements.AddNew(className, classObjectType )
		element.Alias = classAlias
		element.Stereotype = classStereotype
		element.Notes = classNotes
		element.Update
		element.TaggedValues.Refresh
		currentPackage.elements.Refresh
		call LOGInfo( "Added Element: " & element.Name & " (ID=" & element.ElementID & ") in package " & currentPackage.Name )
		set elementTag = nothing
		if not (classC4Type = "") then
			set elementTag = element.TaggedValues.GetByName("Type")
			if elementTag is nothing then
				'add new tag if not found
				set elementTag = element.TaggedValues.AddNew("Type", classTech)
				call LOGInfo( "Add elementTag: " & elementTag.Value & " (ID=" & elementTag.PropertyID & ")")
			else
				elementTag.Value = classTech
				elementTag.Update
				call LOGInfo( "Updated elementTag: " & elementTag.Value & " (ID=" & elementTag.PropertyID & ")")				
			end if
		end if
	else
		set element = Repository.GetElementByID(elements.GetAt(0).ElementID)
	end if

	'add to class_array
	class_array (idxC,0) = element.ElementID
	class_array (idxC,1) = element.Name
	class_array (idxC,2) = element.Alias
	class_array (idxC,3) = element.Stereotype
	class_array (idxC,4) = element.Type
	class_array (idxC,5) = classTech
	if instr(script, "{") > 0 then
		class_array (idxC,6) = "Start"
	end if 
	idxC=idxC+1
	
	'restore logging level	
'	LOGLEVEL=LOGLEVEL_SAVE		

end sub

sub create_sterotype(script)

end sub

sub create_boundary(script)
'		if element.Type = "Boundary" then
'			dim borderStyle
'			set borderStyle = element.Properties("BorderStyle")
'			borderStyle.value = "Solid"
'		end if
	'add to class_array
	class_array (idxC,0) = element.ElementID
	class_array (idxC,1) = element.Name
	class_array (idxC,2) = element.Alias
	class_array (idxC,3) = element.Stereotype	
	class_array (idxC,4) = element.ObjectType
	class_array (idxC,5) = classTech
	if instr(script, "{") > 0 then
		class_array (idxC,6) = "Start"
	end if 
	idxC=idxC+1

end sub

sub end_of(script)
	class_array (idxC,6) = "End"
	idxC=idxC+1
end sub

sub create_relationship(script)
	call LOGTrace("create_relationship()")
	dim strRel
	dim strRelParms
	dim connID
	dim connSterotype
	dim i
	
	'process REL 
	if instr(script,"Rel") > 0 then
		'extract relationship which is delimitered by ()
		strRel = trim(Mid(script, instr(script,"(")+1, instr(script,")")-instr(script,"(")-1))
		call LOGDebug( "strRel: " & strRel )
		strRelParms=split(strRel,",")	
		call LOGDebug( "strRelParms: " & ubound(strRelParms) )
		relationship_array (idxR,0) = getClassID(strRelParms(0))	'source elementID			
		relationship_array (idxR,1) = getClassID(strRelParms(1))	'target elementID
		if ubound(strRelParms) > 1 then
			relationship_array (idxR,3) = strRelParms(2)
			relationship_array (idxR,3) = replace(relationship_array (idxR,3), """", "")
		end if
		if ubound(strRelParms) > 2 then
			connSterotype = strRelParms(3)
			connSterotype = replace(connSterotype, """", "")
		end if
		connID = getConnectorID (relationship_array (idxR,0), relationship_array (idxR,1))
		if connID=0 then
			connID = createConnector (relationship_array (idxR,0), relationship_array (idxR,1), relationship_array (idxR,3), connSterotype)
		end if
		relationship_array (idxR,2) = connID
		idxR=idxR+1
	else
		LOGWarning ( "not yet supported: " & script )
	end if
	
end sub

sub layout_objects()
	call LOGTrace("layout_Objects()")

	dim level	
	level=0
	
	'set the hierarchial levels
	for idxL = 0 to idxC-1	
		if class_array (idxL,6) = "End" then
			level=level-1
		else
			if class_array (idxL,6) = "Start" then
				level=level+1
			end if
		end if
		layout_array (idxL,0) = level
	next

	layout_array (0,1) = 1
	layout_array (0,2) = 1

	'set layout based on level..
	for idxL = 1 to idxC-1	
		if layout_array (idxL,0) = layout_array (idxL-1,0) then
			if class_array (idxL-1,6) = "Start" then
				layout_array (idxL,1) = layout_array (idxL-1,1) + 1
				layout_array (idxL,2) = layout_array (idxL-1,2) 
			else
				layout_array (idxL,1) = layout_array (idxL-1,1)
				layout_array (idxL,2) = layout_array (idxL-1,2) + 1
			end if
		else
			if layout_array (idxL,0) > layout_array (idxL-1,0) then
				layout_array (idxL,1) = layout_array (idxL-1,1) + 1
				layout_array (idxL,2) = layout_array (idxL-1,2) + 1
			else
				layout_array (idxL,1) = layout_array (idxL-1,1) + 1
				layout_array (idxL,2) = layout_array (getParent(idxL),2) 
			end if
		end if
	next

	for idxL = 0 to idxC-1	
		if not(class_array (idxL,6) = "End") then
			layout_array (idxL,3) = getHeight(idxL)
			layout_array (idxL,4) = getWidth(idxL)
		end if
	next

end sub

sub build_diagram()
	call LOGTrace("build_diagram()")
	dim diagramObjects as EA.Collection
	dim diagramObject as EA.DiagramObject
	dim left
	dim right
	dim top
	dim bottom
	dim diagramObjectName
	dim i
	dim width
	dim height
	dim padding
	set DiagramObjects = currentDiagram.DiagramObjects

	width=150
	height=108
	padding=40
	
	' Add diagramObjects
	for i = 0 to idxL-1	
		if not (class_array(i,6) = "End") then
			left = (layout_array (i,2)-1) * (width + padding) + 40
			if layout_array (i,4) = 1 then
				right = left + width
			else
				right = left + ((width + padding) * (layout_array (i,4)+1))
			end if
			top = ((layout_array (i,1)-1) * (height+padding) + 40)*(-1)
			if layout_array (i,3) = 1 then
				bottom = top + height
			else
				bottom = top - ((height + padding) * layout_array (i,3)) 
			end if
			diagramObjectName= "l=" & left & ";r=" & right & ";t=" & top & ";b=" & bottom
			call LOGDebug ( "ElementId=" & class_array (i,0) & "-Left(" & left & ");Right(" & right & ");top(" & top & ");Bottom(" & bottom & ")")
			set diagramObject = currentDiagram.DiagramObjects.AddNew(diagramObjectName, class_array (i,4))
			diagramObject.ElementID = class_array (i,0)
			diagramObject.Update
			diagramObjects.Refresh
		end if
	next

	' Add diagramLinks
	for i = 0 to idxR-1	

	next
	
	currentDiagram.Update

end sub

function getParent(startFrom)
	'call LOGTrace("getParent(" & startFrom & ")")

	'return the index of the immediate parent by going back up the layout_array based on level
	Dim i
	dim level
	
	getParent=startFrom
	level = layout_array (startFrom,0)

	for i = startFrom-1 to 0 step -1
		if layout_array (i,0) = level then
			getParent=i
			exit for
		end if
	next 
	
	'call LOGTrace("getParent=" & getParent)

end function

function getHeight(startFrom)
	call LOGTrace("getHeight(" & startFrom & ")")

	'return the number of rows for this level (including sublevels)
	Dim i
	dim top
	dim bottom
	
	if class_array (startFrom,6) = "Start" then
		top = layout_array (startFrom,1)
		'call LOGDebug ( "top for startFrom(" & startFrom & ")=" & top & " for level=" & layout_array (startFrom,0))
		for i = startFrom+1 to idxC
			'call LOGDebug ( "getheight loop for i= " & i & " of " & idxC & " level=" & layout_array (i,0) & " cont ind=" & class_array (i,6))
			if	class_array (i,6) = "End" then
				if	layout_array (i,0) < layout_array (startFrom,0) then
					bottom=layout_array (i,1)
					'call LOGDebug ( "bottom for i=(" & i & ")=" & bottom)
					getHeight = bottom-top
					exit for
				end if
			end if
		next
	else
		if class_array (startFrom,6) = "" then
			getHeight=1
		end if
	end if
	
	call LOGTrace("getHeight=" & getHeight)

end function

function getWidth(startFrom)
	call LOGTrace("getWidth(" & startFrom & ")")

	'return the max number of cols for a level
	Dim i
	Dim left
	Dim right
	left = layout_array (startFrom,2)
	right = left+1
	'call LOGDebug ( "left for startFrom(" & startFrom & ")=" & left & " for level=" & layout_array (startFrom,0))

	if class_array (startFrom,6) = "Start" then
		for i = startFrom to idxC
			'call LOGDebug ( "getwidth loop for i= " & i & " of " & idxC & " level=" & layout_array (i,0) & " cont ind=" & class_array (i,6))
			if layout_array (i,2) > right then		
				right = layout_array (i,2)
			end if
			if class_array (i,6) = "End" then
				if layout_array (i,0) < layout_array (startFrom,0) then
					exit for
				end if
			end if
			'call LOGDebug ( "getwidth loop for i= " & i & " of " & idxC & " left=" & left & " right=" & right)			
		next 
	end if
	getWidth = right - left
	call LOGTrace("getWidth=" & getWidth)

end function

function getClassID(name)
	call LOGTrace("getClassID(" & name & ")")
	
	dim i
	
	getClassID=99
	for i = 0 to idxC
		if name = class_array(i,3) then				'check using alias
			getClassID = class_array (i,0)
			exit for
		else
			if name = class_array(i,2) then			'check using name
				getClassID = class_array (i,0)
				exit for
			end if
		end if
	next
	if getClassID = 99 then
		LOGWarning ( "class not found: " & name )
	end if
	
	call LOGTrace("getClassID=" & getClassID )
	
end function
	
function getConnectorID(fromElementID, toElementID)
	call LOGTrace("getConnectorID(" & fromElementID & "," & toElementID & ")")

	dim element as EA.Element
	dim connector as EA.Connector

	getConnectorID=0
	set element = Repository.GetElementByID (fromElementID)
	for each connector in element.Connectors
		if connector.SupplierID = toElementID then
			getConnectorID=connector.ConnectorID
			exit for
		end if
	next
	call LOGTrace("getConnectorID=" & getConnectorID )

end function

function createConnector(fromElementID, toElementID, connName, connSterotype)
	call LOGTrace("createConnector(" & fromElementID & "," & toElementID & "," & connName & ")")
	dim element as EA.Element
	dim connector as EA.Connector

	set element = Repository.GetElementByID (fromElementID)
	
	'create connector
	set connector = element.Connectors.AddNew("","Connector")	
	connector.SupplierID = toElementID
	connector.Name = connName
	connector.Direction = "Source -> Destination"
	connector.Stereotype = connSterotype
	connector.Update
	element.Connectors.Refresh		
	createConnector = connector.ConnectorID
	call LOGTrace("createConnector=" & createConnector )
		
end function