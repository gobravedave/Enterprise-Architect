'
' Script Name: Create C4 PlantUML
' Author: David Anderson
' Purpose: Callable routine for dealing with the creation of a C4 PlantUML script
' Date: 18-Sept-2022
'
dim generated_script
'dim reltag_array (99,4)
' 0=sterotype
' 1=
'dim iRT

sub CreateC4PlantUML ()
	call LOGInfo("Create C4 PlantUML script activated")

	dim diagramObject as EA.DiagramObject
	dim element as EA.Element
	dim connector as EA.Connector
	dim elementTag as EA.TaggedValue

	dim include_puml
	
	select case currentDiagram.Type
		case "Logical"		include_puml = "C4_Context"
		case "Deployment"	include_puml = "C4_Container"
		case "Component"	include_puml = "C4_Component"
	end select

	generated_script="@startuml" & vbcrlf & _
	"!include https://raw.githubusercontent.com/plantuml-stdlib/C4-PlantUML/master/" & include_puml & ".puml" & vbcrlf & _
	"' uncomment the following line and comment the first to use locally" & vbcrlf & _
	"' !include " & include_puml & ".puml" & vbcrlf & _
	vbcrlf & _
	"'LAYOUT_TOP_DOWN()" & vbcrlf & _
	"'LAYOUT_AS_SKETCH()" & vbcrlf & _
	"LAYOUT_WITH_LEGEND()" & vbcrlf

	generated_script = generated_script & "title " & currentDiagram.Name & vbcrlf

	for each diagramObject in currentDiagram.DiagramObjects
		set Element = Repository.GetElementByID(diagramObject.ElementID)
		if not(theSelectedElement.ElementID = element.ElementID) then		
			element_array (e,0) = element.ElementID
			element_array (e,1) = element.Type
			if instr(element.Name, " ") = 0 then
				element_array (e,2) = element.Name
			else
				element_array (e,2) = chr(34) & element.Name & chr(34)
			end if
			'replace spaces with \n if length greater than 20 
			'if len(element_array (e,2)) > 20 then
			'	element_array (e,2) = replace(element_array (e,2), " ","\n")
			'end if
			if element.Alias = "" then
				element_array (e,3) = replace(element.Name, " ", "_")
			else
				if instr(element.Alias, " ") = 0 then
					element_array (e,3) = element.Alias	
				else
					element_array (e,3) = chr(34) & element.Alias & chr(34)
				end if
			end if
			'boundary objects do not have a sterotype type value
			if element.Stereotype = "" then
				if element.Type = "Boundary" then
					element_array (e,4) =  "Boundary"
				end if
			else
				element_array (e,4) = replace(element.Stereotype, " ", "_")
			end if
			element_array (e,5) = element.Notes

			if element.TaggedValues.Count > 0 then
				set elementTag = element.TaggedValues.GetByName("Type")
				if elementTag is nothing then
					call LOGWarning("Element: " & element.Name & " missing 'Type' tagged value")
				else
					element_array (e,6) = elementTag.Value
				end if
			else
'				element_array (e,6) = lcase(color(element.Type, diagramObject.BackgroundColor))
				element_array (e,6)=""
			end if
			element_array (e,7) = diagramObject.left	
			element_array (e,8) = diagramObject.top	
			element_array (e,9) = diagramObject.right
			element_array (e,10) = diagramObject.bottom
			element_array (e,11) = ""			'children - delimitered string of element ids within a boundary
			element_array (e,12) = ""			'parent boundary elementid			
			e=e+1							
		'	resolve tags		
		end if
	next
	
	'resolve boundary hierarchy using coordinates
	dim i
	dim i2
	dim i2Id
	For i = 0 to e-1
		if element_array (i,1) = "Boundary" then		
			'loop thru the element array and check is the element coordinates are inside the boundary
			'call LOGDebug("*process boundary(" & i & "):" & element_array(i,0))
			For i2 = 0 to e-1
				if not(i = i2) then			
					if element_array (i2,1) = "Boundary" then
						if inBoundary(i2,i) then
						'add boundary element to the list of children elements for this boundary
							i2Id = "B" & i2 & "#" & element_array (i2,0)
							if element_array (i,11) = "" then
								element_array (i,11) = i2Id
							else
								element_array (i,11) = element_array (i,11) & "|" & i2Id 
							end if
							'update the element with the boundary parent.. 
							if element_array (i2,12) = "" then
								element_array (i2,12) = element_array (i,0) 
							else
								element_array (i2,12) = element_array (i2,12) & "|" & element_array (i,0) 
							end if
						end if
					end if
				end if
			Next 
		end if
	Next 

	'call LOGDebug("Find Children: part 1")

'	for each boundary.. populate the elements within
	dim children
	dim i3		'index for the boundary in the element array
	dim i4		'index to compare 
	for i = 0 to e-1
		if element_array (i,1) = "Boundary" then
			if not element_array (i,11)="" then
				'call LOGDebug("*children:" & i & "=" & element_array (i,11))
				children = split( element_array (i,11),"|")
				for i2 = ubound(children) to 0 step -1
				'start with the deepest boundary and work out..
					'call LOGDebug("*processing boundary childs:" & i2 & "=" & children(i2))
					'check if this is a boundary...
					if left(children(i2),1)="B" then
						'get index of the boundary we are checking against..
						'remove hardcoded length
						i3 = Mid(children(i2),2,1)
						'call LOGDebug ( "index: " & i3)					
						'loop thru the elements arrary
						'compare coordinate
						For i4 = 0 to e-1
							if not (element_array (i4,1) = "Boundary") then
								if inBoundary(i4,i3) then
									if element_array (i3,11) = "" then
										element_array (i3,11) = element_array (i4,0)
									else
										element_array (i3,11) = element_array (i3,11) & "|" & element_array (i4,0) 
									end if
									'update the element with the boundary parent.. 
									if element_array (i4,12) = "" then
										element_array (i4,12) = element_array (i3,0) 
									else
										element_array (i4,12) = element_array (i4,12) & "|" & element_array (i3,0) 
									end if
								end if
							end if
						next
					end if
				next
			end if
		end if
	next 
	
	'for each top level boundary.. populate the elements within if they have not yet been allocated

	'call LOGDebug("Find Children: part 2")
	for i = 0 to e-1
		if element_array (i,1) = "Boundary" then
			if element_array (i,12)="" then
				'call LOGDebug("Find Children: part 2" & i & "=" & element_array (i,0))
				For i4 = 0 to e-1
					if not (element_array (i4,1) = "Boundary") then
						if element_array (i4,12) = "" then
							if inBoundary(i4,i) then
								if element_array (i,11) = "" then
									element_array (i,11) = element_array (i4,0)
								else
									element_array (i,11) = element_array (i,11) & "|" & element_array (i4,0) 
								end if
								'update the element with the boundary parent.. 
								if element_array (i4,12) = "" then
									element_array (i4,12) = element_array (i,0) 
								else
									element_array (i4,12) = element_array (i4,12) & "|" & element_array (i,0) 
								end if
							end if
						end if
					end if
				next
			end if
		end if
	next 
				
	call LOGDebug( "Element Array" )
	call PrintArray (element_array,0,e-1)

	'spin through diagram links
	dim diagramLink as EA.DiagramLink
	r=0
	for each diagramLink in currentDiagram.DiagramLinks
		set connector = Repository.GetConnectorByID(diagramLink.ConnectorID)
		relationship_array (r,0) = connector.ClientID
		relationship_array (r,1) = connector.SupplierID
		relationship_array (r,2) = connector.Name
		relationship_array (r,3) = connector.Stereotype
		'TODO: build up relTag for each sterotype and set color and  linetype
		r=r+1
	next

	dim strLine
	'Output PlantUML objects
	generated_script = generated_script & "AddElementTag('Service', $shape=EightSidedShape(), $legendText='service (eight sided)')" & vbcrlf
	generated_script = generated_script & "AddRelTag(Alert, $textColor='red', $lineColor='red', $lineStyle = DashedLine(), $legendText='Alert')" & vbcrlf
	generated_script = generated_script & "AddRelTag(Synchronous, $textColor='black', $lineColor='dimgrey',$legendText='Synchronous')" & vbcrlf
	generated_script = generated_script & "AddRelTag(Asynchronous,$textColor='black', $lineColor='dimgrey', $lineStyle = DashedLine(), $legendText='Asynchronous')" & vbcrlf
	
	'Process top level objects (without a parent)
	for i = 0 to e-1
		if element_array (i,12) = "" then
			strLine = construct_output(i)
			'No Children exist
			if element_array (i,11) = "" then
				generated_script = generated_script & strLine & vbcrlf
			else
				generated_script = generated_script & strLine & "{" & vbcrlf
				'Process Boundary
				call process_boundary (i)
				generated_script = generated_script & "}" & vbcrlf
			end if
		end if
	next 
	
	'Output PlantUML relationships
	for i = 0 to r-1
		strLine = "Rel(" & system(relationship_array (i,0)) & "," & _		
					system(relationship_array (i,1)) & "," & _
					chr(34) & relationship_array (i,2) & chr(34) & "," & _
					chr(34) & relationship_array (i,3) & chr(34)
		if not (relationship_array (i,3)="") then
			strLine = strLine & ", $tags='" & relationship_array (i,3) & "'"
		end if
		strLine = strLine & ")"
		generated_script = generated_script & strLine & vbcrlf
	next 
	
	generated_script=generated_script & "SHOW_LEGEND()" & vbcrlf & "@enduml" & vbcrlf
	theSelectedElement.Notes = generated_script
	theSelectedElement.Update

end sub

sub process_boundary (i)
	call LOGDebug("Process Boundary:" & element_array (i,11))
	dim strLine
	dim member
	dim i2
	dim i3
	dim i4
	member = split(element_array (i,11),"|")
	for i2 = 0 to ubound(member)
		call LOGDebug("Process Member:" & member(i2))
		'check if this is a boundary...
		if left(member(i2),1)="B" then
			'get index of the boundary 
			i3 = Mid(member(i2),2,1)
			strLine = construct_output(i3)
			if element_array (i3,11) = "" then
				generated_script = generated_script & strLine & vbcrlf
			else
				generated_script = generated_script & strLine & "{" & vbcrlf
				'Process Boundary
				call process_boundary (i3)
				generated_script = generated_script & "}" & vbcrlf
			end if
		else
			'get index based using element id
			i4 = getIndex(CLng(member(i2)))
			strLine = construct_output(i4)
			generated_script = generated_script & strLine & vbcrlf
		end if
	next
	
end sub

function construct_output (i)
	'call LOGDebug ( "construct_output: " & i & "=" & element_array (i,0))		

	dim strLine
	dim strType
	'if boundary..
	if inStr("...Person,System...", element_array (i,4)) > 0 then
		strLine = element_array (i,4) & "(" & _		
				element_array (i,3) & ", " & _		
				element_array (i,2) 
	else
		if element_array (i,1) = "Boundary" then
			strLine = element_array (i,4) & "(" & _		
					element_array (i,3) & ", " & _		
					element_array (i,2) 
		else
			if element_array (i,6) = "Storage" then
				strType=element_array (i,4) & "Db"
			else
				strType=element_array (i,4)
			end if
			strLine = strType & "(" & _		
					element_array (i,3) & ", " & _		
					element_array (i,2) & ", " & _
					element_array (i,6)
			if element_array (i,6) = "Service" then
				strLine = strLine & ", $tags='Service'"
			'else
			'	strLine = strLine & ","
			end if		
		end if
	end if
	if element_array (i,5) = "" then
		strLine = strLine & ")"		
	else
		strLine = strLine & "," & chr(34) & element_array (i,5) & chr(34) & ")"
	end if
	construct_output = strLine
end function

function inBoundary(i1, i2) 
	'call LOGDebug ( "inBoundary: " & i1 & ":" & i2)					

	if element_array (i1,7) > element_array (i2,7) _			
		and element_array (i1,8) < element_array (i2,8) _		
		and element_array (i1,9) < element_array (i2,9) _		
		and element_array (i1,10) > element_array (i2,10) then
		inBoundary=True
	else
		inBoundary=False
	end if
	'call LOGDebug ( "inBoundary=" & inBoundary)					

end function

function system(elementid)
	Dim i
	'Call LOGTrace( "system(" & elementid & ")" )

	For i = 0 to Ubound(element_array)
		if element_array (i,0) = "" then
			Exit for
		end if
		if element_array (i,0) = elementid then
			system = element_array (i,3)		'return alias
			Exit for
		end if
	Next
	'Call LOGTrace( "system=" & system )

end function

function getIndex(elementid)
	Dim i
	'Call LOGTrace( "getIndex(" & elementid & "of type " & TypeName(elementid) & ")" )

	For i = 0 to Ubound(element_array)
		'Call LOGTrace( "i(" & i & ")=" & element_array (i,0) & " of type=" & TypeName(element_array (i,0)) )
		if element_array (i,0) = "" then
			Exit for
		end if
		if element_array (i,0) = elementid then
			getIndex = i		'return index
			Exit for
		end if
	Next
	'Call LOGTrace( "getIndex=" & getIndex )

end function