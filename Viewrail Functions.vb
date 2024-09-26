Imports System.Runtime.CompilerServices
Imports Inventor
Imports Autodesk.iLogic.Automation
Imports Autodesk.iLogic.Runtime

'Namespace Doc
Public Module DocumentExtensions

        Private ThisApplication As Object = GetObject(, "Inventor.Application")
        Private vaultApp As ApplicationAddIn = ThisApplication.ApplicationAddIns.ItemById("{48B682BC-42E6-4953-84C5-3D253B52E77B}")

        ''' <summary>
        ''' Activates give view
        ''' </summary>
        ''' <param name="doc">Assembly Document</param>
        ''' <param name="viewName">View Name</param>
        <Extension()>
        Sub ActivateView(ByRef doc As Document, viewName As String)

            Dim docViews As DesignViewRepresentations = doc.ComponentDefinition.RepresentationsManager.DesignViewRepresentations
            Dim createView As Boolean = True
            
            For Each v As DesignViewRepresentation In docViews
                If v.Name = viewName Then
                    v.Activate
                    v.Locked = False
                    createView = False
                    Exit For
                End If
            Next v
            
            If createView Then
                docViews.Add(viewName)
                docViews.Item(viewName).Activate
            End If
            
        End Sub 'ActivateView


        ''' <summary>
        ''' Clears the contents of the Windows Clipboard.
        ''' </summary>
        Sub ClipboardClear()

            My.Computer.Clipboard.Clear()

        End Sub


        ''' <summary>
        ''' Returns the contents of the Windows Clipboard as a string.
        ''' </summary>
        ''' <returns>String</returns>
        Function ClipboardGet() As String
	
            Return My.Computer.Clipboard.GetText()
            
        End Function 'ClipboardGet

        
        ''' <summary>
        ''' Sets the contents of the Windows Clipboard to the given string value.
        ''' </summary>
        ''' <param name="clipboardText">Text to store in the Clipboard.</param>
        Sub ClipboardSet(clipboardText As String)

            If clipboardText Is Nothing Or clipboardText = "" Then
                My.Computer.Clipboard.Clear()
            Else            
                My.Computer.Clipboard.SetText(clipboardText)
            End If
            
        End Sub 'ClipboardSet


        ''' <summary>
        ''' Returns False if given component has an unconstrained sketch
        ''' </summary>
        ''' <param name="comp">Component Occurrence</param>
        ''' <returns>Boolean</returns>
        <Extension()>
        Function CompHasBadSketches(iLogicVb As ILowLevelSupport, comp As ComponentOccurrence) As Boolean
	
            Dim badSketches As Boolean = False
        
            If Not comp.Suppressed Then
                For Each s As PlanarSketch In comp.Definition.Sketches
                    If s.ConstraintStatus = 51714 Then
                        iLogicVb.Automation.LogControl.Log(LogLevel.Info, s.Name & " is not a fully constrained sketch in " & comp.Name & " (" & s.ConstraintStatus & ")")
                        badSketches = True            
                    End If
                Next s
            End If
                                        
            Return badSketches
            
        End Function 'CompHasBadSketches


        ''' <summary>
        ''' Given Document Object and Component Name returns True/False if component exists in Document
        ''' </summary>
        ''' <param name="doc">Document Object</param>
        ''' <param name="paramName">Component Name</param>
        ''' <returns>Boolean</returns>
        <Extension()>
        Public Function ComponentExists(doc As Document, compName As String) As Boolean

            For Each comp As ComponentOccurrence In doc.ComponentDefinition.Occurrences
                If comp.Name = compName Then Return True
            Next comp
            Return False
            
        End Function 'ComponentExists
		
		
        ''' <summary>
        ''' Checks if a constraint exists in a given AssemblyDocument.
        ''' </summary>
        ''' <param name="doc">AssemblyDocument to be checked</param>
        ''' <param name="conName">Name of constraint being checked for</param>
        Function ConstraintExists(doc As Document, conName As String)
            For Each con As AssemblyConstraint In doc.ComponentDefinition.Constraints
                If con.Name = conName Then Return True
            Next con
            Return False
        End Function 'ConstraintExists()


        ''' <summary>
        ''' Copies a file from one location to another.
        ''' </summary>
        ''' <param name="sourceFile">Full path to the source file.</param>
        ''' <param name="targetFile">Full path to the target file.</param>
        Sub CopyFile(sourceFile As String, targetFile As String)

            If Not System.IO.File.Exists(targetFile) Then System.IO.File.Copy(sourceFile, targetFile, True)
            My.Computer.FileSystem.GetFileInfo(targetFile).IsReadOnly = False
            
        End Sub 'CopyFile
		

        ''' <summary>
        ''' Copies Parameter values from sourceDoc to targetDoc. It will only copy parameters that have identical names in both documents.
	''' Optional list of parameter (names) to skip and dictionary of differently names parameter to push
        ''' </summary>
        ''' <param name="sourceDoc">Source Document Object</param>
        ''' <param name="targetDoc">Target Document Object</param>
        ''' <param name="skipParams">Parameters to Skip</param>
        ''' <param name="mapParams">Dictionary of parameter map</param>
        Sub CopyParameters(sourceDoc As Document, targetDoc As Document, Optional skipParams As List(Of String) = Nothing, Optional mapParams As Dictionary(Of String, String) = Nothing)

            Dim alwaysSkipParams As New List(Of String) From { "debug" }

            If skipParams Is Nothing Then
                skipParams = alwaysSkipParams
            Else
                skipParams.AddRange(alwaysSkipParams)
            End If
                
            For Each param As UserParameter In sourceDoc.ComponentDefinition.Parameters.UserParameters
                If ParameterExists(targetDoc, param.Name) AndAlso Not skipParams.Contains(param.Name) Then
                    targetDoc.ComponentDefinition.Parameters(param.Name).Value = param.Value
                End If
            Next param
            
            If Not mapParams Is Nothing Then
                For Each kvp As KeyValuePair(Of String, String) In mapParams
                    If ParameterExists(sourceDoc, kvp.Key) AndAlso ParameterExists(targetDoc, kvp.Value) Then
                        targetDoc.ComponentDefinition.Parameters(kvp.Value).Value = sourceDoc.ComponentDefinition.Parameters.UserParameters(kvp.Key).Value
                    End If
                Next kvp
            End If

        End Sub 'CopyParameters


        ''' <summary>
        ''' Exports specified face as a DXF file.
        ''' </summary>
        ''' <param name="comp">ComponentOccurrence to export from</param>
        ''' <param name="faceName">Name of face to export</param>
        ''' <param name="fileName">Name of exported DXF file</param>
        Sub ExportDXF(comp As ComponentOccurrence, faceName As String, fileName As String)

            Dim compDef As PartComponentDefinition = comp.Definition
            Dim compDoc As Document = compDef.Document
            Dim face As FaceProxy = GetFaceProxy(comp, faceName)
            Dim sketch As PlanarSketch = compDef.Sketches.Add(face.NativeObject)
            
            For Each el As EdgeLoop In face.NativeObject.EdgeLoops
                For Each e As Edge In el.Edges
                    Dim ent As SketchEntity = sketch.AddByProjectingEntity(E)
                    ent.Construction = False
                Next e
            Next el
            
            sketch.DataIO.WriteDataToFile("DXF", fileName)
            sketch.Delete()
            
        End Sub 'ExportDXF
		

        ''' <summary>
        ''' Given Component Occurrence and Feature Name returns True/False if Feature exists in Component
        ''' </summary>
        ''' <param name="comp">Component Occurrence</param>
        ''' <param name="featureName">Feature Name</param>
        ''' <returns></returns>
        <Extension()>
        Function FeatureExists(comp As ComponentOccurrence, featureName As String) As Boolean

            Return FeatureExists(comp.Definition.Document, featureName)
            
        End Function 'FeatureExists
		

        ''' <summary>
        ''' Given Component Occurrence and Feature Name returns True/False if Feature exists in Document
        ''' </summary>
        ''' <param name="doc">Part Document</param>
        ''' <param name="featureName">Feature Name</param>
        ''' <returns></returns>
        <Extension()>
        Function FeatureExists(doc As Document, featureName As String) As Boolean

            For Each f As PartFeature In doc.ComponentDefinition.Features
                If f.Name = featureName Then Return True
            Next f
            
            Return False
            
        End Function 'FeatureExists


        ''' <summary>
        ''' Formats a phone number by adding parentheses, spaces, and dashes.
        ''' </summary>
        ''' <param name="phoneNumber">The phone number to be formatted.</param>
        ''' <returns>The formatted phone number.</returns>
        Function FormatPhoneNumber(phoneNumber As String) As String

            Dim formattedPhoneNumber As String = Nothing
            Dim phoneNum As String = System.Text.RegularExpressions.Regex.Replace(phoneNumber, "\D", "")
            
            If phoneNum.Length = 10 Then
                formattedPhoneNumber = "(" & phoneNum.Substring(0, 3) & ") " & phoneNum.Substring(3, 3) & "-" & phoneNum.Substring(6, 4)
            ElseIf phoneNum.Length = 11 Then
                formattedPhoneNumber = phoneNum.Substring(0, 1) & " (" & phoneNum.Substring(1, 3) & ") " & phoneNum.Substring(4, 3) & "-" & phoneNum.Substring(7, 4)
            Else
                formattedPhoneNumber = phoneNum
            End If
            
            Return formattedPhoneNumber
            
        End Function 'FormatPhoneNumber
        

        ''' <summary>
        ''' Returns a ComponentOccurrence object based upon a user selection.
        ''' </summary>
        ''' <param name="prompt">Prompt string</param>
        ''' <returns>ComponentOccurrence</returns>
        Function GetComponent(prompt As String) As ComponentOccurrence

            Dim entity As Object = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, prompt)
           
            Return entity

        End Function 'GetComponent
		
		
        ''' <summary>
        ''' Returns distance between two parallel objects in inches
        ''' </summary>
        ''' <param name="startObj">Start Object</param>
        ''' <param name="endObj">End Object</param>
        ''' <returns>Double</returns>
        Function GetDistance(startObj As Object, endObj As Object) As Double

            Dim startPoint As Object = Nothing
            Dim endPoint As Point = Nothing

           Select Case True
				Case TypeName(startObj).Contains("WorkPoint")
					startPoint = startObj.Point
				Case TypeName(startObj).Contains("WorkAxis")
					startPoint = startObj.Line
				Case TypeName(startObj).Contains("WorkPlane")
					startPoint = startObj.Plane
				Case TypeName(startObj).Contains("Face")
					startPoint = startObj.Geometry
               Case Else
                   Throw New Exception("Unknown startObj type: (" & startObj.Type)
			End Select

            endPoint = GetThePoint(endObj)
            If endPoint Is Nothing Then Throw New Exception("Unknown endObj type: " & endObj.Type)
            
            Dim distance As Double = Nothing
            distance = Abs(startPoint.DistanceTo(endPoint)) / 2.54
            
            Return distance

        End Function 'GetDistance


        ''' <summary>
        ''' Calculates the distance between two points on a specified plane.
        ''' </summary>
        ''' <param name="startObj">The starting point object.</param>
        ''' <param name="endObj">The ending point object.</param>
        ''' <param name="axisName">The name of the axis (X, Y, or Z).</param>
        ''' <returns>The distance between the two points on the specified plane, in inches.</returns>
        Function GetDistanceOnAxis(startObj As Object, endObj As Object, axisName As String) As Double

            Dim startPoint As Point = GetThePoint(startObj)
            Dim endPoint As Point = GetThePoint(endObj)
            
            If startPoint Is Nothing Then Throw New Exception("Cannot get a distance on plane from " & TypeName(startObj))
            If endPoint Is Nothing Then Throw New Exception("Cannot get a distance on plane from " & TypeName(endObj))
                
            Dim distance As Double = 0	
            Dim pointToPoint As Vector = startPoint.VectorTo(endPoint)
            
            Select Case axisName
                Case "X"
                    distance = pointToPoint.X
                Case "Y"
                    distance = pointToPoint.Y
                Case "Z"
                    distance = pointToPoint.Z
            End Select
            
            Return Abs(distance / 2.54)

        End Function 'GetDistanceOnAxis


        ''' <summary>
        ''' Prompt user to select the face of an object, returns FaceProxy of their selection.
        ''' </summary>
        ''' <param name="prompt">Prompt string</param>
        ''' <returns>FaceProxy</returns>
        Function GetFace(prompt As String) As FaceProxy
        
            Dim clickedFace As FaceProxy = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kPartFaceFilter, prompt)
            
            Return clickedFace
            
        End Function 'GetFace
        

        ''' <summary>
        ''' Retrieves the Document / FactoryDocument object associated with the specified baseObject.
        ''' </summary>
        ''' <param name="baseObject">The baseObject for which to retrieve the Document.</param>
        ''' <returns>The Document object associated with the baseObject.</returns>
        Function GetDocument(baseObject As Object) As Document
            
            Dim returnDocument As Document

            Select Case TypeName(baseObject)
                
                Case "CadDoc", "AssemblyComponentDefinition", "PartComponentDefinition"
                    returnDocument = baseObject.Document
                    
                Case "_DocumentClass", "AssemblyDocument", "PartDocument", "DrawingDocument"
                    returnDocument = baseObject
                    
                Case "ComponentOccurrence", "ComponentOccurrenceProxy"
                    returnDocument = baseObject.Definition.Document
                    
                Case Else
                    Throw New Exception("Object type " & TypeName(baseObject) & " not accounted for in GetDocument Function")
                
            End Select
            
            If returnDocument.DocumentType <> Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
                If returnDocument.ComponentDefinition.IsModelStateMember Then
                    returnDocument = returnDocument.ComponentDefinition.FactoryDocument
                End If
            End If
            
            Return returnDocument

        End Function 'GetDocument


        ''' <summary>
        ''' Returns name of given face
        ''' </summary>
        ''' <param name="faceObject">faceObject</param>
        Function GetFaceName(faceObject As Object) As String
	
            Dim pickFace As Face
            
            Select Case TypeName(faceObject)
                Case "FaceProxy"
                    pickFace = faceObject.NativeObject
                Case "Face"
                    pickFace = faceObject
                Case Else
                    Throw New Exception("Invalid Face type: " & TypeName(faceObject))
            End Select
            
            Dim faceName As String = Nothing
            Dim attSets As AttributeSets = pickFace.AttributeSets
            
            If attSets.NameIsUsed("iLogicEntityNameSet") Then
                Dim attSet As AttributeSet = attSets.Item("iLogicEntityNameSet")		
                For Each at As Attribute In attSet
                    If at.Name = "iLogicEntityName" Then
                        faceName = at.Value
                        Exit For
                    End If
                Next at
            End If
            
            Return faceName
            
        End Function 'GetFaceName


        ''' <summary>
        ''' Given component occurrence and face name, returns face proxy if it exists
        ''' </summary>
        ''' <param name="comp">Component Occurrence</param>
        ''' <param name="faceName">Face Name</param>
        ''' <returns>FaceProxy</returns>
        <Extension()>
        Public Function GetFaceProxy(comp As ComponentOccurrence, faceName As String) As FaceProxy
            
            For Each bod As SurfaceBody In comp.SurfaceBodies
                For Each faceProx As FaceProxy In bod.Faces
                    Dim attSets = faceProx.NativeObject.AttributeSets
                    If attSets.NameIsUsed("iLogicEntityNameSet") Then
                        Dim attSet As AttributeSet = attSets.Item("iLogicEntityNameSet")
                        For Each at As Attribute In attSet
                            If at.Name = "iLogicEntityName" Then
                                If at.Value = faceName Then
                                    Return faceProx
                                End If
                            End If
                        Next at
                    End If
                Next faceProx
            Next bod
            
            Return Nothing
            
        End Function 'GetFaceProxy


        ''' <summary>
        ''' Given component occurrence and face name, returns face proxy if it exists
        ''' </summary>
        ''' <param name="comp">Component Occurrence</param>
        ''' <param name="faceName">Face Name</param>
        ''' <returns>FaceProxy</returns>
        <Extension()>
        Public Function GetFaceProxy(comp As ComponentOccurrenceProxy, faceName As String) As FaceProxy
            
            Return GetFaceProxy(CType(comp, ComponentOccurrence), faceName)
            
        End Function 'GetFaceProxy


        ''' <summary>
        ''' Retrieves the file name of the specified component occurrence.
        ''' </summary>
        ''' <param name="comp">The component occurrence.</param>
        ''' <returns>The file name of the component occurrence, or nothing if the component or its referenced file descriptor is nothing.</returns>
        <Extension()>
        Function GetFileName(comp As ComponentOccurrence) As String

            Dim fileName As String = Nothing

            If comp IsNot Nothing AndAlso comp.ReferencedFileDescriptor IsNot Nothing Then
                fileName = comp.ReferencedFileDescriptor.FullFileName
            End If

            Return fileName

        End Function 'GetFileName


        ''' <summary>
        ''' Returns Leaf Occurrence of selected object.
        ''' </summary>
        ''' <param name="prompt">Prompt for object to select</param>
        ''' <returns>ComponentOccurrence</returns>
        Function GetLeafComponent(prompt As String) As ComponentOccurrence

            Dim entity As Object = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAssemblyLeafOccurrenceFilter, prompt)
           
            Return entity

        End Function 'GetLeafComponent

        
        ''' <summary>
        ''' Returns the part number associated with the given component.
        ''' </summary>
        ''' <param name="comp">ComponentOccurrence</param>
        ''' <returns>String</returns>
        <Extension()>
        Public Function GetPartNumber(comp As ComponentOccurrence) As String

            Dim partNumber As String = Nothing

            If TypeOf comp.Definition Is VirtualComponentDefinition Then 
                partNumber = comp.Definition.PropertySets("Design Tracking Properties")("Part Number").Value
            Else 
                partNumber = GetDocument(comp).PropertySets("Design Tracking Properties")("Part Number").Value
            End If

            Return partNumber

        End Function 'GetPartNumber


        ''' <summary>
        ''' Returns a top level proxy of the given thing in the given component occurrence
        ''' </summary>
        ''' <param name="comp">Component Occurrence</param>
        ''' <param name="thing">Object to create proxy from</param>
        ''' <returns>Object Proxy</returns>
        Function GetProxy(comp As ComponentOccurrence, thing As Object) As Object

            Dim parentComp As ComponentOccurrence = comp.ParentOccurrence
            Dim curProxy As Object
            Dim prevProxy As Object
            
            comp.CreateGeometryProxy(thing, curProxy)
            
            While parentComp IsNot Nothing
                prevProxy = curProxy
                parentComp.CreateGeometryProxy(prevProxy, curProxy)
                parentComp = parentComp.ParentOccurrence
            End While
            
            Return curProxy

        End Function 'GetProxy


        ''' <summary>
        ''' Returns a SurfaceBody based upon a user selected component.
        ''' </summary>
        ''' <param name="prompt">Selection Prompt</param>
        ''' <param name="name"></param>
        ''' <returns>SurfaceBody</returns>
        Function GetSurfaceBody(prompt As String) As SurfaceBody
        
            Dim clickedBody As SurfaceBody = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kPartBodyFilter, prompt)
            
            Return clickedBody
            
        End Function 'GetSurfaceBody


        ''' <summary>
        ''' Returns the root point of a given base object.
        ''' </summary>
        ''' <param name="baseObject">The object from which to retrieve the root point.</param>
        ''' <returns>The root point of the given object.</returns>
        Function GetThePoint(baseObject As Object) As Point
        
            Dim thePoint As Point = Nothing
            
            Select Case True
            Case TypeName(baseObject).Contains("WorkPoint")
                thePoint = baseObject.Point
            Case TypeName(baseObject).Contains("WorkAxis")
                thePoint = baseObject.Line.RootPoint
            Case TypeName(baseObject).Contains("WorkPlane")
                thePoint = baseObject.Plane.RootPoint
            Case TypeName(baseObject).Contains("Face")
                thePoint = baseObject.Geometry.RootPoint
            Case TypeName(baseObject).Contains("Cylinder")
                thePoint = baseObject.BasePoint
            End Select
            
            Return thePoint
            
        End Function 'GetThePoint


        ''' <summary>
        ''' Returns the top level parent occurrence of a given component occurrence
        ''' </summary>
        ''' <param name="comp">Component Occurrence</param>
        ''' <returns>Component Occurrence</returns>
        <Extension()>
        Function Grandparent(comp As ComponentOccurrence) As ComponentOccurrence
	
            If comp.ParentOccurrence Is Nothing Then Return Nothing
            
            While comp.ParentOccurrence IsNot Nothing
                comp = comp.ParentOccurrence
            End While
            
            Return comp
            
        End Function 'Grandparent


        ''' <summary>
        ''' Returns True if all given assembly document has no broken constraints
        ''' </summary>
        ''' <param name="doc">Assembly Document</param>
        ''' <returns>Boolean</returns>
        <Extension()>
        Function HasBrokenConstraints(doc As Document, Optional iLogicVb As ILowLevelSupport = Nothing) As Boolean
	
            Dim broken = False
            
            For Each c As AssemblyConstraint In doc.ComponentDefinition.Constraints
                If Not {"kRedundantHealth", "kSuppressedHealth", "kUpToDateHealth"}.Contains(c.HealthStatus.ToString) Then
                    broken = True
                    If iLogicVb Is Nothing Then
                        Exit For
                    Else
                        iLogicVb.Automation.LogControl.Log(LogLevel.Error, "Broken Constraint: " & c.Name)
                    End If
                End If
            Next c
            
            Return broken
            
        End Function 'HasBrokenConstraints


        ''' <summary>
        ''' Returns True if given component occurrence is fully constrained (except for rotation)
        ''' </summary>
        ''' <param name="comp">Component Occurrence</param>
        ''' <returns>Boolean</returns>
        <Extension()>
        Function IsFullyConstrained(comp As ComponentOccurrence) As Boolean
	
            Dim translationDegreesCount As Integer
            Dim translationDegreesVectors As ObjectsEnumerator
            Dim rotationDegreesCount As Integer
            Dim rotationDegreesVectors As ObjectsEnumerator
            Dim dofCenter As Point
            Dim constrained As Boolean = False
            
            comp.GetDegreesOfFreedom(translationDegreesCount, translationDegreesVectors, rotationDegreesCount, rotationDegreesVectors, dofCenter)
            If translationDegreesCount = 0 And rotationDegreesCount = 0 Then constrained = True
                
            Return constrained
            
        End Function 'IsFullyConstrained


        ''' <summary>
        ''' Returns True if local files have been purged today
        ''' </summary>
        ''' <returns>Boolean</returns>
        Function LocalFolderCheck() as Boolean
            Dim OK As Boolean = True
            Dim verifiedFolders As Integer = 0
            Dim today As String = DateTime.Now.ToShortDateString
            Dim dirs As System.IO.DirectoryInfo() = New System.IO.DirectoryInfo("C:\$WorkingFolder\Engineering Data").GetDirectories
            
            For Each d As System.IO.DirectoryInfo In dirs
                If d.CreationTime.Date.ToShortDateString = today Then
                    verifiedFolders += 1
                Else
                    OK = False
                End If
            Next d
            
            Return OK
            
        End Function 'LocalFolderCheck


        ''' <summary>
        ''' Given Document Object and Parameter names returns True if parameter exists in Document
        ''' </summary>
        ''' <param name="doc">Document Object</param>
        ''' <param name="paramName">Parameter Name</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ParameterExists(doc As Document, paramName As String) As Boolean

            Dim params As Parameters
            If doc.DocumentType = Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
                params = doc.Parameters
            Else
                params = doc.ComponentDefinition.Parameters
            End If
        
            For Each p As Parameter In params
                If p.Name = paramName Then Return True
            Next p

            Return False

        End Function 'ParameterExists


        ''' <summary>
        ''' Given Document Object and Pattern name returns True/False if pattern exists in Document
        ''' </summary>
        ''' <param name="doc">Document Object</param>
        ''' <param name="patternName">Pattern Name</param>
        ''' <returns></returns>
        ''' <remarks>Do not use in Drawing Documents</remarks>
        <Extension()>
        Public Function PatternExists(doc As Document, patternName As String) As Boolean
        
            For Each p As OccurrencePattern In doc.ComponentDefinition.OccurrencePatterns
                If p.Name = patternName Then Return True
            Next p
            
            Return False

        End Function 'PatternExists
		
		
        ''' <summary>
        ''' Calculates the relative position between two component occurrences along a specified axis.
        ''' Measurement is the 2nd component relavtive to the 1st.
        ''' Derek, make sure both components are in the same assembly
        ''' </summary>
        ''' <param name="comp1">The first component occurrence.</param>
        ''' <param name="comp2">The second component occurrence.</param>
        ''' <param name="axisName">The name of the axis ("X", "Y", or "Z").</param>
        ''' <returns>The relative position between the two component occurrences along the specified axis, in inches.</returns>
        Function RelativePosition(comp1 As ComponentOccurrence, comp2 As ComponentOccurrence, axisName As String) As Double

            Dim offset As Double

            Select Case axisName
                Case "X"
                    offset = comp2.Transformation.Translation.X - comp1.Transformation.Translation.X
                Case "Y"
                    offset = comp2.Transformation.Translation.Y - comp1.Transformation.Translation.Y
                Case "Z"
                    offset = comp2.Transformation.Translation.Z - comp1.Transformation.Translation.Z
            End Select

            Return offset / 2.54

        End Function


        ''' <summary>
        ''' Returns a rounded value based on the specified precision.
        ''' </summary>
        ''' <param name="num">Double</param>
        ''' <param name="precision">Double</param>
        ''' <returns>Double</returns>
        Function RoundTo(num As Double, precision As Double) As Double
	
            Return(Round(num / precision) * precision)
            
        End Function 'RoundTo


        ''' <summary>
        ''' Sets the current view port to the "home" view (either default or user configured)
        ''' </summary>
        Sub SetHomeView()
	
            ThisApplication.CommandManager.ControlDefinitions.Item("AppViewCubeHomeCmd").Execute
            ThisApplication.ActiveView.Fit
            
        End Sub 'SetHomeView


        ''' <summary>
        ''' Assigns the specified partNumber to the listed component.
        ''' </summary>
        ''' <param name="comp">ComponentOccurrence</param>
        ''' <param name="partNumber">String</param>
        <Extension()>
        Sub SetPartNumber(comp As ComponentOccurrence, partNumber As String)

            GetDocument(comp).PropertySets.Item("Design Tracking Properties").Item("Part Number").Value = partNumber

        End Sub 'SetPartNumber


        ''' <summary>
        ''' Toggles a component using the Component wrapper. If the component is in a pattern, then entire pattern is toggled
        ''' </summary>
        ''' <param name="iLogicVb">iLogic Object</param>
        ''' <param name="comp">Component Occurrence</param>
        ''' <param name="active">State</param>
        Sub ToggleComp(iLogicVb As ILowLevelSupport, comp As ComponentOccurrence, active As Boolean)
	
            If comp Is Nothing Then Throw New Exception("Not Implemented")
            Dim compName As String = comp.Name
            
            If comp.IsPatternElement Then
                If comp.PatternElement.Parent.IsPatternElement Then
                    compName = comp.PatternElement.Parent.PatternElement.Parent.Name
                Else
                    compName = comp.PatternElement.Parent.Name
                End If
            End If
            
            If comp.ParentOccurrence Is Nothing Then
                iLogicVb.CreateObjectProvider(GetDocument(iLogicVb.RuleDocument)).Component.IsActive(compName) = active
            Else
                iLogicVb.CreateObjectProvider(GetDocument(comp.ParentOccurrence)).Component.IsActive(compName) = active
            End If
            
        End Sub 'ToggleComp


        ''' <summary>
        ''' Causes the computer to speak the given text string.
        ''' </summary>
        ''' <param name="textToSpeech">String</param>
        Sub TTS(textToSpeech As String)

            If textToSpeech Is Nothing Then Exit Sub
            If textToSpeech = "" Then Exit Sub

            Dim spVoice As Object = CreateObject("SAPI.SpVoice")
            spVoice.Speak(textToSpeech)

        End Sub 'TTS


        ''' <summary>
        ''' Removes all constraints from the given component (name)
        ''' </summary>
        ''' <param name="iLogicVb">iLogic Wrapper</param>
        ''' <param name="partName">Component Name</param>
        ''' <param name="active">Suppression State</param>
        <Extension()>
        Sub Unconstrain(iLogicVb As ILowLevelSupport, partName As String, Optional active As Boolean = True)
	
            Dim iComp As ICadComponent = iLogicVb.CreateObjectProvider(GetDocument(iLogicVb.RuleDocument)).Component
            Dim comp As ComponentOccurrence = iComp.InventorComponent(partName)
            Dim partConstraints As AssemblyConstraintsEnumerator = comp.Constraints

            Try
                If Not comp.Suppressed Then comp.Grounded = False
            Catch
                iLogicVb.Automation.LogControl.Log(LogLevel.Error, "Unabled to automatically unGround " & partName)
            End Try

            For Each c As AssemblyConstraint In partConstraints
                c.Delete()
            Next c

            iComp.IsActive(partName) = active

        End Sub 'Unconstrain


        ''' <summary>
        ''' Returns true if the named work plane exists in the specified document.
        ''' </summary>
        ''' <param name="doc">Document</param>
        ''' <param name="wpName">String</param>
        ''' <returns>Boolean</returns>
        Public Function WorkPlaneExists(doc As Document, wpName As String) As Boolean

            For Each wp As WorkPlane In doc.ComponentDefinition.WorkPlanes
                If wp.Name = wpName Then Return True
            Next wp
            
            Return False

        End Function


        ''' <summary>
        ''' Writes the given data to the specified file name.
        ''' The file is created if it does not exist and will be overwritten if it does exist.
        ''' </summary>
        ''' <param name="fileName">String</param>
        ''' <param name="data">String</param>
        Sub WriteTextToFile(fileName As String, data As String)
	
            Dim oWrite As System.IO.StreamWriter = System.IO.File.CreateText(fileName)
            oWrite.WriteLine(data)
            oWrite.Close()
            
        End Sub 'WriteTextToFile


End Module
'End Namespace
