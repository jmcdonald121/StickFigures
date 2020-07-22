VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub StackMapTextLabels()
    'This will label every selected feature in the polygon coverage
    'or all features if no features are selected
    'This was formerly known as LabelUsingRelatedField()
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pLayer As ILayer
    Dim pFlayer As IFeatureLayer
    Dim pFClass As IFeatureClass
    Dim pFieldName As String
    Dim pFieldName2 As String
    Dim pQFilt As IQueryFilter
    Dim pFeature As IFeature
    Dim pFeatCursor As IFeatureCursor
    Dim pFeatureSelection As IFeatureSelection
    Dim pSelectionSet As ISelectionSet
    Dim pIndex As Long
    Dim pFeatCount As Long
      
    Set pMxDoc = ThisDocument
    Set pMap = pMxDoc.FocusMap
    
    pFieldName = "GEO_TYPE_CODE" 'Enter field name here
    pFieldName2 = "GEO_DEPTH_TO"
  
    'Must have the feature class added to ArcMap as the first layer
    If pMap.LayerCount < 1 Then
        MsgBox "No layer in TOC. Please add a polygon layer.", _
        vbOKOnly, "Missing layer"
        Exit Sub
    End If
  
    Set pLayer = pMap.Layer(0)
    Set pFlayer = pLayer
    
    'Open the FeatureClass
    Set pFClass = pFlayer.FeatureClass
    
    'Labels all features with each features related rows
    'Get any selected features
    Set pFeatureSelection = pFlayer 'QI
    Set pSelectionSet = pFeatureSelection.SelectionSet
    'Create an empty query filter
    Set pQFilt = New QueryFilter
    pQFilt.WhereClause = ""
    If pSelectionSet.Count < 1 Then
        MsgBox "Please select one or more wells", vbCritical, "No Selected Wells"
        Exit Sub
        'Retrieve all features
        'Set pFeatCursor = pFClass.Search(pQFilt, False)
    Else
        'Retrieve just the selected features
        'This statement moves the selected features into a feature cursor.  All
        'subsequent work is done using the feature cursor.
        pSelectionSet.Search pQFilt, False, pFeatCursor
    End If
    
    'Get the first feature
    Set pFeature = pFeatCursor.NextFeature
    
    'Get the first Relationship Class
    Dim pEnumRelClass As IEnumRelationshipClass
    Dim pRelationshipClass As IRelationshipClass
    
    Set pEnumRelClass = pFeature.Class.RelationshipClasses(esriRelRoleAny)
    Set pRelationshipClass = pEnumRelClass.Next
  
    If pRelationshipClass Is Nothing Then
        MsgBox "     No Relationship Class created for this layer." & _
        vbNewLine & "Please create a Relationship Class in ArcCatalog", _
        vbOKOnly, "Relationship Class Missing"
        Exit Sub
    End If
   
    'Loop through the features
    Do
        'Create a set that contains this object
        Dim pFeatSet As ISet
        Set pFeatSet = New esriSystem.Set
        pFeatSet.Add pFeature
        
        'Now loop through using GetObjectsRelatedToObject
        Dim i As Long
        Dim r As Long
        Dim myCount As Integer
        Dim pRelObjSet As ISet
        Dim pFieldIndex As Double
        Dim pFieldIndex2 As Double
        Dim pObjectClass As IObjectClass 'Holds the destination table
        
        Dim pRelRow As IRow
        Set pRelObjSet = pRelationshipClass.GetObjectsRelatedToObject(pFeature)
        myCount = pRelObjSet.Count
        
        Dim myArray() As String
        ReDim myArray(pRelObjSet.Count - 1)
        
        Dim myArray2() As Double
        ReDim myArray2(pRelObjSet.Count - 1)
    
        'Get the destination field
        Set pObjectClass = pRelationshipClass.DestinationClass
        pFieldIndex = pObjectClass.FindField(pFieldName)
        pFieldIndex2 = pObjectClass.FindField(pFieldName2)
        
        For i = 0 To myCount - 1
            Set pRelRow = pRelObjSet.Next
            'Get the value of the field you want to label with
            'and use it to create a label expression
            If VarType(pRelRow.Value(pFieldIndex)) = vbNull Then
                myArray(i) = "-999"
            Else
                myArray(i) = pRelRow.Value(pFieldIndex)
            End If
            myArray2(i) = pRelRow.Value(pFieldIndex2)
        Next i
        
        'Perform a bubble sort on the data
        Dim AnyChanges As Boolean
        Dim SwapFH As Variant
        Dim SwapFH2 As Variant
        Do
            AnyChanges = False
            For i = LBound(myArray2) To UBound(myArray2) - 1
                If (myArray2(i) > myArray2(i + 1)) Then
                ' These two need to be swapped
                SwapFH = myArray2(i)
                SwapFH2 = myArray(i)
                myArray2(i) = myArray2(i + 1)
                myArray(i) = myArray(i + 1)
                myArray2(i + 1) = SwapFH
                myArray(i + 1) = SwapFH2
                AnyChanges = True
                End If
            Next i
        Loop Until Not AnyChanges 'This is the end of the bubble sort
    
        'Build label expression
        Dim strExp As String
        Dim strExp1 As String
            For i = 0 To myCount - 1
            If Not strExp = " " Then
                strExp = strExp & vbNewLine & myArray(i) & ", " & myArray2(i)
            Else
                strExp = myArray(i)
            End If
        Next i
'        Debug.Print strExp
        Call AddLabel(strExp, pFeature)
        
        'Clear the label string
        strExp = ""
        
        Set pFeature = pFeatCursor.NextFeature
    Loop Until pFeature Is Nothing
  
End Sub

Private Sub AddLabel(strLabel As String, pFeature As IFeature)
    Dim pMxDoc As IMxDocument
    Dim pGraphicsContainer As IGraphicsContainer
    Dim pActiveView As IActiveView
    Dim pTextElement As ITextElement
    Dim pElement As IElement
    Dim pPoint As IPoint
    
    Set pMxDoc = Application.Document
    Set pGraphicsContainer = pMxDoc.FocusMap
    Set pActiveView = pMxDoc.FocusMap
    
    Set pTextElement = New TextElement
    Set pElement = pTextElement
    
    'Get the point location
    Set pPoint = pFeature.Shape
    pTextElement.Text = strLabel 'Add the text string to the Text Symbol
    
    Dim myFont As IFontDisp
    Set myFont = New StdFont
    myFont.Name = "Courier New"
    myFont.size = 1

    pTextElement.Symbol.size = 5#   'Added Text Symbol Size, 20060623, Jim McDonald
    pTextElement.Symbol.Font = myFont
    pElement.Geometry = pPoint
    pGraphicsContainer.AddElement pTextElement, 0
    
    pActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
    
End Sub

Public Sub DeleteAllGraphics()
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pGraCon As IGraphicsContainer
    Dim pAV As IActiveView
    
    Set pMxDoc = ThisDocument
    Set pMap = pMxDoc.FocusMap
    Set pAV = pMap
    Set pGraCon = pMap
    
    pGraCon.DeleteAllElements
    
    pAV.Refresh
    
End Sub

Public Sub StackMapStickFigures()
    'This will label every selected feature in the polygon coverage
    'or all features if no features are selected
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pLayer As ILayer
    Dim pFlayer As IFeatureLayer
    Dim pFClass As IFeatureClass
    Dim pFieldName As String
    Dim pFieldName2 As String
    Dim pQFilt As IQueryFilter
    Dim pFeature As IFeature
    Dim pFeatCursor As IFeatureCursor
    Dim pFeatureSelection As IFeatureSelection
    Dim pSelectionSet As ISelectionSet
    Dim pIndex As Long
    Dim pFeatCount As Long
      
    Set pMxDoc = ThisDocument
    Set pMap = pMxDoc.FocusMap
    
    pFieldName = "INDICT" 'Enter field name here
    pFieldName2 = "GEO_DEPTH_TO"
  
    'Must have the feature class added to ArcMap as the first layer
    If pMap.LayerCount < 1 Then
        MsgBox "No layer in TOC. Please add a polygon layer.", _
        vbOKOnly, "Missing layer"
        Exit Sub
    End If
  
    Set pLayer = pMap.Layer(0)
    Set pFlayer = pLayer
    
    'Open the FeatureClass
    Set pFClass = pFlayer.FeatureClass
    
    'Labels all features with each features related rows
    'Get any selected features
    Set pFeatureSelection = pFlayer 'QI
    Set pSelectionSet = pFeatureSelection.SelectionSet
    'Create an empty query filter
    Set pQFilt = New QueryFilter
    pQFilt.WhereClause = ""
    If pSelectionSet.Count < 1 Then
        MsgBox "Please select one or more wells", vbCritical, "No Selected Wells"
        Exit Sub
        'Retrieve all features
        'Set pFeatCursor = pFClass.Search(pQFilt, False)
    Else
        'Retrieve just the selected features
        'This statement moves the selected features into a feature cursor.  All
        'subsequent work is done using the feature cursor.
        pSelectionSet.Search pQFilt, False, pFeatCursor
    End If
    
    'Get the first feature
    Set pFeature = pFeatCursor.NextFeature
    
    'Get the first Relationship Class
    Dim pEnumRelClass As IEnumRelationshipClass
    Dim pRelationshipClass As IRelationshipClass
    
    Set pEnumRelClass = pFeature.Class.RelationshipClasses(esriRelRoleAny)
    Set pRelationshipClass = pEnumRelClass.Next
  
    If pRelationshipClass Is Nothing Then
        MsgBox "     No Relationship Class created for this layer." & _
        vbNewLine & "Please create a Relationship Class in ArcCatalog", _
        vbOKOnly, "Relationship Class Missing"
        Exit Sub
    End If
   
    'Loop through the features
    Do
        'Create a set that contains this object
        Dim pFeatSet As ISet
        Set pFeatSet = New esriSystem.Set
        pFeatSet.Add pFeature
        
        'Now loop through using GetObjectsRelatedToObject
        Dim i As Long
        Dim r As Long
        Dim myCount As Integer
        Dim pRelObjSet As ISet
        Dim pFieldIndex As Double
        Dim pFieldIndex2 As Double
        Dim pObjectClass As IObjectClass 'Holds the destination table
        
        Dim pRelRow As IRow
        Set pRelObjSet = pRelationshipClass.GetObjectsRelatedToObject(pFeature)
        myCount = pRelObjSet.Count
        
        Dim myArray() As String
        ReDim myArray(pRelObjSet.Count - 1)
        
        Dim myArray2() As Double
        ReDim myArray2(pRelObjSet.Count - 1)
    
        'Get the destination field
        Set pObjectClass = pRelationshipClass.DestinationClass
        pFieldIndex = pObjectClass.FindField(pFieldName)
        pFieldIndex2 = pObjectClass.FindField(pFieldName2)
        
        For i = 0 To myCount - 1
            Set pRelRow = pRelObjSet.Next
            'Get the value of the field you want to label with
            'and use it to create a label expression
            If VarType(pRelRow.Value(pFieldIndex)) = vbNull Then
                myArray(i) = "-999"
            Else
                myArray(i) = pRelRow.Value(pFieldIndex)
            End If
            myArray2(i) = pRelRow.Value(pFieldIndex2)
        Next i
        
        'Perform a bubble sort on the data
        Dim AnyChanges As Boolean
        Dim SwapFH As Variant
        Dim SwapFH2 As Variant
        Do
            AnyChanges = False
            For i = LBound(myArray2) To UBound(myArray2) - 1
                If (myArray2(i) > myArray2(i + 1)) Then
                ' These two need to be swapped
                SwapFH = myArray2(i)
                SwapFH2 = myArray(i)
                myArray2(i) = myArray2(i + 1)
                myArray(i) = myArray(i + 1)
                myArray2(i + 1) = SwapFH
                myArray(i + 1) = SwapFH2
                AnyChanges = True
                End If
            Next i
        Loop Until Not AnyChanges 'This is the end of the bubble sort
    
        'Build label expression
        Dim strExp As String
        Dim strExp1 As String
            For i = 0 To myCount - 1
            If Not strExp = " " Then
                strExp = strExp & vbNewLine & myArray(i) & ", " & myArray2(i)
            Else
                strExp = myArray(i)
            End If
        Next i
'        Debug.Print strExp 'This is used to debug the code and data
        Call AddStick(myCount, myArray, myArray2, pFeature)
        
        'Clear the label string
        strExp = ""
        
        Set pFeature = pFeatCursor.NextFeature
    Loop Until pFeature Is Nothing
  
End Sub

Private Function ColorPolyonGraphics(txtComp As String) As ISimpleFillSymbol
    Dim pMxDoc As IMxDocument
    Dim pStyleGall As IStyleGallery
    Dim pEnumFills As IEnumStyleGalleryItem
    Dim pStyleGallItem As IStyleGalleryItem
    Dim pSimpleFillSym As IFillSymbol
    Dim pLineSymbol As ISimpleLineSymbol
    Dim pColor As IColor
    Dim pLineColor As IColor
            
    Set pMxDoc = ThisDocument
    Set pStyleGall = pMxDoc.StyleGallery
    Set pEnumFills = pStyleGall.Items("Colors", "ESRI.style", "Default")
    Set pStyleGallItem = pEnumFills.Next
    
    Set pSimpleFillSym = New SimpleFillSymbol
    Set pLineSymbol = New SimpleLineSymbol
    
    Do Until pStyleGallItem Is Nothing
        Select Case txtComp
        Case "1" 'Clay
            If pStyleGallItem.Name = "Cretean Blue" Then
                Set pColor = pStyleGallItem.Item
            End If
        Case "2" 'Silt
            If pStyleGallItem.Name = "Apatite Blue" Then Set pColor = pStyleGallItem.Item
        Case "3" 'Sand
            If pStyleGallItem.Name = "Solar Yellow" Then Set pColor = pStyleGallItem.Item
        Case "4" 'Gravel
            If pStyleGallItem.Name = "Cantaloupe" Then Set pColor = pStyleGallItem.Item
        Case "5" 'Diamict
            If pStyleGallItem.Name = "Peridot Green" Then Set pColor = pStyleGallItem.Item
        Case "6" 'Rock
            If pStyleGallItem.Name = "Gray 30%" Then Set pColor = pStyleGallItem.Item
        Case Else 'Not Classified
            If pStyleGallItem.Name = "Black" Then Set pColor = pStyleGallItem.Item
        End Select
        Set pStyleGallItem = pEnumFills.Next
    Loop
    
    'Set up the outline color, width, and style
    Set pLineColor = New RgbColor
    pLineColor.RGB = vbBlack
    pLineSymbol.Color = pLineColor
    pLineSymbol.Width = 0.01
    pLineSymbol.Style = esriSLSSolid
    
    'Finish setting up the Polygon symbol, i.e. color and outline
    pSimpleFillSym.Color = pColor
    pSimpleFillSym.Outline = pLineSymbol
    
    Set ColorPolyonGraphics = pSimpleFillSym
    
End Function

Private Sub AddStick(intStackCount As Integer, myArray() As String, myArray2() As Double, pFeature As IFeature)
    Dim pMxDoc As IMxDocument
    Dim pGraphicsContainer As IGraphicsContainer
    Dim pActiveView As IActiveView
    Dim pTextElement As ITextElement
    Dim pPoint As IPoint
    Dim i As Integer, j As Integer
    Dim dblX As Double
    Dim dblY As Double
    Dim dblDelX As Double
    Dim dblDelY As Double
    Dim dblThk As Double
    Dim pNewPt As IPoint
    Dim pPoly As IPointCollection
'    Dim pGeom As IGeometry
    Dim pElement As IElement
    Dim pPolygonElem As IPolygonElement
    Dim pFillShapeElem As IFillShapeElement
    Dim pGroupElem As IGroupElement
    'These variable are used to only display the first encountered bedrock unit.
    Dim blnBedrock As Boolean
    
    'Start
    Set pMxDoc = Application.Document
    Set pGraphicsContainer = pMxDoc.FocusMap
    Set pActiveView = pMxDoc.FocusMap
    
    'Get the point location
    Set pPoint = pFeature.Shape
    dblX = pPoint.X
    dblY = pPoint.Y
    'Create the width of the box.  We can set this to a user defined width later.
    dblDelX = dblX + 100# '100# is the width of the stack figure in map feet.
    dblDelY = dblY
    dblThk = 0
    
    'Create polygon graphic boxes
    blnBedrock = False
    Set pGroupElem = New GroupElement
    For i = 0 To intStackCount - 1
    
        'This code should only display the first bedrock unit - Jim McDonald, 20061208
        If blnBedrock = True Then Exit For
        If myArray(i) = 6 Then blnBedrock = True
        'End of display of first bedrock unit code - Jim McDonald, 20061208
        
        Set pPoly = New Polygon
        
        Select Case i
            Case 0
                dblThk = myArray2(i) * 10 '10 is the vertical exaggeration
            Case Else
                dblThk = (myArray2(i) - myArray2(i - 1)) * 10 '10 is the vertical exaggeration
        End Select
        
        For j = 0 To 4
            Set pNewPt = New Point
            
            Select Case j
            Case 0, 4
                pNewPt.PutCoords dblX, dblDelY
            Case 1
                pNewPt.PutCoords dblDelX, dblDelY
            Case 2
                pNewPt.PutCoords dblDelX, dblDelY - dblThk
            Case 3
                pNewPt.PutCoords dblX, dblDelY - dblThk
            End Select
            
            pPoly.AddPoint pNewPt
        Next j
        
        dblDelY = dblDelY - dblThk
        
        Debug.Print dblThk
        
        Set pPolygonElem = New PolygonElement
        Set pFillShapeElem = pPolygonElem
        pFillShapeElem.Symbol = ColorPolyonGraphics(myArray(i))
        
        Set pElement = pPolygonElem
        pElement.Geometry = pPoly
        
        pGroupElem.AddElement pElement
        
    Next i
    
    pGraphicsContainer.AddElement pGroupElem, 0
    
    pActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
    
End Sub

Private Function ColorPolyonGraphics_Old(txtComp As String) As IFillSymbol
    Dim pMxDoc As IMxDocument
    Dim pStyleGall As IStyleGallery
    Dim pEnumFills As IEnumStyleGalleryItem
    Dim pFill As IStyleGalleryItem
    Dim pSimpleFill As IFillSymbol
            
    Set pMxDoc = ThisDocument
    Set pStyleGall = pMxDoc.StyleGallery
    Set pEnumFills = pStyleGall.Items("Fill Symbols", "ESRI.style", "Default")
    Set pFill = pEnumFills.Next
    Do Until pFill Is Nothing
        Select Case txtComp
        Case "1" 'Clay
            If pFill.Name = "Med Blue" Then Set pSimpleFill = pFill.Item
        Case "2" 'Silt
            If pFill.Name = "Blue" Then Set pSimpleFill = pFill.Item
        Case "3" 'Sand
            If pFill.Name = "Yellow" Then Set pSimpleFill = pFill.Item
        Case "4" 'Gravel
            If pFill.Name = "Orange" Then Set pSimpleFill = pFill.Item
        Case "5" 'Diamict
            If pFill.Name = "Med Green" Then Set pSimpleFill = pFill.Item
        Case "6" 'Rock
            If pFill.Name = "Gray 30%" Then Set pSimpleFill = pFill.Item
        Case Else 'Not Classified
            If pFill.Name = "Biohazard Overlay" Then Set pSimpleFill = pFill.Item
        End Select
        Set pFill = pEnumFills.Next
    Loop
    
    Set ColorPolyonGraphics = pSimpleFill
    
End Function

Public Sub WWTopCalc()
    'This macro calculates the tops for the water wells, using the bottoms supplied in the table.
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pEnumLyr As IEnumLayer
    Dim pLayer As ILayer
    Dim pFlayer As IFeatureLayer
    Dim pFC As IFeatureClass
    Dim lngNumFeatures As Long
    Dim pFCur1 As IFeatureCursor
    Dim pFCur2 As IFeatureCursor
    Dim pQF1 As IQueryFilter
    Dim pQF2 As IQueryFilter
    Dim pFeature1 As IFeature
    Dim pFeature2 As IFeature
    Dim lngFldGeoTop As Long
    Dim lngFldGeoDepth As Long
    Dim lngFldWellNo As Long
    Dim lngFldFID As Long
    Dim i As Long
    Dim dblDepth As Double
    Dim dblTop As Double
    
    'Start
    Set pMxDoc = ThisDocument
    Set pMap = pMxDoc.FocusMap
    
    'Find the layer which will have the tops calculated.
    Set pEnumLyr = pMap.Layers
    Set pLayer = pEnumLyr.Next
    Do Until pLayer Is Nothing
        If pLayer.Name = "sortnoH2O" Then 'Currently, this points to a shape file.  But
            Set pFlayer = pLayer
            Set pFC = pFlayer.FeatureClass 'Identify the featureclass of the required layer
                                            'NOTE:  This feature class has already been presorted.
                                            'In the future, we can do a ITableSort instead of presorting the table.
        End If
        Set pLayer = pEnumLyr.Next
    Loop
    
    'Find the relevent fields in the featureclass
    lngFldGeoDepth = pFC.FindField("GEO_DEPTH_")
    lngFldGeoTop = pFC.FindField("GEO_TOP")
    lngFldWellNo = pFC.FindField("WELL_LOG_N")
    lngFldFID = pFC.FindField("FID") 'FID only applies to Shape files.  For Personnel Geodatabases, use OBJECTID instead.
    'Count the number of features
    lngNumFeatures = pFC.FeatureCount(Nothing)
    
    'Create the new query filters
    Set pQF1 = New QueryFilter
    Set pQF2 = New QueryFilter
        
    For i = 0 To lngNumFeatures - 1 'Cycle through all the features
        If i = 0 Then
            pQF2.WhereClause = "FID = " & i
            Set pFCur2 = pFC.Update(pQF2, False)
            Set pFeature2 = pFCur2.NextFeature
            pFeature2.Value(lngFldGeoTop) = 0
        ElseIf i > 0 And i <= lngNumFeatures - 1 Then
            pQF2.WhereClause = "FID = " & i
            Set pFCur2 = pFC.Update(pQF2, False)
            Set pFeature2 = pFCur2.NextFeature
            
            pQF1.WhereClause = "FID = " & i - 1
            Set pFCur1 = pFC.Update(pQF1, False)
            Set pFeature1 = pFCur1.NextFeature
            
            If pFeature1.Value(lngFldWellNo) = pFeature2.Value(lngFldWellNo) Then
                pFeature2.Value(lngFldGeoTop) = pFeature1.Value(lngFldGeoDepth)
            Else
                pFeature2.Value(lngFldGeoTop) = 0
            End If
        End If
        
        pFCur2.UpdateFeature pFeature2 'Add/Update the new values back into the table
    Next
    
End Sub

