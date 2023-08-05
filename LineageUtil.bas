Attribute VB_Name = "LineageUtil"
Sub FindShapeByID()

    ' Declare variables
    Dim shp As Shape
    Dim id As Long
    Dim found As Boolean

    ' Prompt the user for the shape ID
    id = InputBox("Enter the shape ID:")

    ' Initialize the found flag
    found = False

    ' Loop through all shapes in the active page
    For Each shp In ActivePage.Shapes
        ' Check if the shape ID matches the user input
        If shp.id = id Then
            ' Select the shape
            ActiveWindow.Select shp, visSelect
            ' Set the found flag
            found = True
            ' Exit the loop
            Exit For
        End If
    Next shp

    ' If the shape was not found, display a message
    If Not found Then
        MsgBox "No shape with ID " & id & " was found.", vbInformation
    End If

End Sub

