Attribute VB_Name = "OneNoteSort"


Sub AlphabetizeOneNote_NoteBook()

'Make sure to add the references to:
'Microsoft OneNote 15.0 Object Library
'Microsoft XML, 6.0

    'I put this code together to solve a problem, I realize it is not perfectly optimized.
    'Add this code to your vba editor in Excel and make sure the notebook you want to sort is open in OneNote.
    'If you have any input on how optimize the code, feel free to let me know.
    
Dim notebookSortName As String

'Change the name of the library to the one that needs to alphabetized
notebookSortName = "Put name of notebook here"

Dim onenote As onenote.Application
Set onenote = New onenote.Application

' Get the XML that represents the OneNote sections
Dim oneNoteSectionsXml As String

onenote.GetHierarchy "", 4, oneNoteSectionsXml

Dim doc As MSXML2.DOMDocument
Set doc = New MSXML2.DOMDocument

If doc.LoadXML(oneNoteSectionsXml) Then
    Dim nodeNoteBooks As MSXML2.IXMLDOMNodeList
    Dim nodeSections As MSXML2.IXMLDOMNodeList
    Set nodeNoteBooks = doc.DocumentElement.SelectNodes("//one:Notebook")
    Set nodeSections = doc.DocumentElement.SelectNodes("//one:Section")
    Dim nodeNoteBook As MSXML2.IXMLDOMNode
    Dim nodeSection As MSXML2.IXMLDOMNode
    
    Dim UpdateHierarchyHeader As String
    Dim UpdateHierarchyBody As String
    Dim UpdateHierarchyFooter As String
    
    Dim pageName As String
    Dim notebookID As String
    Dim notebookName As String
    Dim notebookNameXML As String
    Dim notebookIDXML As String
    Dim notebookPathXML As String
    Dim sectionName As String
    Dim sectionNameXML As String
    Dim sectionIDXML As String
    Dim sectionPathXML As String
    Dim parentID As String
    
    Dim i As Integer
    Dim r As Integer
    
    Dim notebookArray() As Variant
    Dim sectionArray() As Variant
    Dim sectionArraySorted() As Variant
    
    
    Dim w As Worksheet 'temp worksheet to transpose and sort section names
    Dim sectionRange As Range
    
    i = 0
    For Each nodeNoteBook In nodeNoteBooks
    i = i + 1
    Next
    
    ReDim notebookArray(0 To i, 0 To 4) As Variant
    
    i = 0
    'put notebook id and names in array
    For Each nodeNoteBook In nodeNoteBooks
    'ReDim test(0 To i, 0 To 1) As Variant
       notebookID = nodeNoteBook.Attributes.getNamedItem("ID").Text
       notebookName = nodeNoteBook.Attributes.getNamedItem("name").Text
       notebookIDXML = nodeNoteBook.Attributes.getNamedItem("ID").XML
       notebookNameXML = nodeNoteBook.Attributes.getNamedItem("name").XML
       notebookPathXML = nodeNoteBook.Attributes.getNamedItem("path").XML
       notebookArray(i, 0) = notebookID
       notebookArray(i, 1) = notebookName
       notebookArray(i, 2) = notebookNameXML
       notebookArray(i, 3) = notebookIDXML
       notebookArray(i, 4) = notebookPathXML
       
       i = i + 1
    Next
    
    ReDim sectionArray(0 To 3, 0)
    r = 0
    For Each nodeSection In nodeSections
    Call onenote.GetHierarchyParent(nodeSection.Attributes.getNamedItem("ID").Text, parentID)
       
        For i = 0 To UBound(notebookArray, 1)
            If notebookArray(i, 0) = parentID Then
                notebookID = notebookArray(i, 0)
                notebookName = notebookArray(i, 1)
                notebookNameXML = notebookArray(i, 2)
                notebookIDXML = notebookArray(i, 3)
                notebookPathXML = notebookArray(i, 4)
            End If
        Next
        
        If notebookName = notebookSortName Then
        ReDim Preserve sectionArray(3, r) As Variant
            sectionIDXML = nodeSection.Attributes.getNamedItem("ID").XML
            sectionNameXML = nodeSection.Attributes.getNamedItem("name").XML
            sectionName = nodeSection.Attributes.getNamedItem("name").Text
            sectionPathXML = nodeSection.Attributes.getNamedItem("path").XML
            

            sectionArray(0, r) = sectionNameXML
            sectionArray(1, r) = sectionName
            sectionArray(2, r) = sectionIDXML
            sectionArray(3, r) = sectionPathXML
            
            'Header for the XML code
            UpdateHierarchyHeader = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>" & vbCrLf & _
                                    "<one:Notebooks xmlns:one=" & Chr(34) & "http://schemas.microsoft.com/office/onenote/2013/onenote" & Chr(34) & ">" & vbCrLf & _
                                    "<one:Notebook " & notebookIDXML & " " & notebookPathXML & ">"
            r = r + 1
        End If
    Next
Else
    MsgBox "OneNote 2013 XML Data failed to load."
End If

'    Transpose and sort array
Set w = ThisWorkbook.Worksheets.Add()
'MsgBox UBound(sectionArray, 2)
Set sectionRange = w.Range(Cells(1, 1), Cells(UBound(sectionArray, 2) + 1, 4))
sectionRange = Application.Transpose(sectionArray)

w.Sort.SortFields.Add Key:=Range(Cells(1, 2), Cells(UBound(sectionArray, 2) + 1, 2)), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With w.Sort
    .SetRange sectionRange
    .Header = xlGuess
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'resize and repopulate sectionArray with the newly transposed data
sectionArray = sectionRange


Application.DisplayAlerts = False
w.Delete    'delete the temp worksheet
Application.DisplayAlerts = True

'populate body of xml
For i = 1 To UBound(sectionArray)
    sectionIDXML = sectionArray(i, 3)  'ID XML
    sectionPathXML = sectionArray(i, 4)  'Path XML
    UpdateHierarchyBody = UpdateHierarchyBody & vbCrLf & "<one:Section " & sectionIDXML & " " & sectionPathXML & "/>"
Next

'populate footer of xml
UpdateHierarchyFooter = "    </one:Notebook>" & vbCrLf & _
                        "</one:Notebooks>"


'combine xml elements and load them into xml reader. Then call OneNote API to update the hierarchy of the notebook
If doc.LoadXML(UpdateHierarchyHeader & UpdateHierarchyBody & UpdateHierarchyFooter) Then
    onenote.UpdateHierarchy doc.XML
Else
    MsgBox "Make sure notebook name is correct."
End If

End Sub
