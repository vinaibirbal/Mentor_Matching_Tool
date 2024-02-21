Attribute VB_Name = "Module3"

Sub Match_Matrix()

Dim wb As Workbook
Dim sh1 As Worksheet
Dim sh2 As Worksheet
Dim sh3 As Worksheet
Dim sh4 As Worksheet
Dim sh5 As Worksheet
Dim sh6 As Worksheet




Set wb = ActiveWorkbook
Set sh1 = wb.Worksheets("Mentees")
Set sh2 = wb.Worksheets("Mentors")
Set sh3 = wb.Worksheets("Weight Matrix")
Set sh4 = wb.Worksheets("Category Weight Values")
Set sh5 = wb.Worksheets("Match")
Set sh6 = wb.Worksheets("mentors_used")

Call Create_Table







Dim lngcolum As Integer
Dim lngrows As Integer

Dim c As Range
Dim c1 As Integer





Dim Lng As Integer ' number of categories used
Lng = sh4.Cells(1, Columns.Count).End(xlToLeft).Column
With sh3 ' populate weight matrix by iterating through matrix
    lngcolumn = .Cells(1, Columns.Count).End(xlToLeft).Column - 2
    lngrows = .Cells(Rows.Count, 2).End(xlUp).Row - 1
    For I = 1 To lngrows
         For J = 1 To lngcolumn
            Call Pair_Value(sh3.Cells(I + 1, 2).Value(), sh3.Cells(1, J + 2).Value(), Lng, I + 1, J + 2)
        Next J
        
    Next I

End With

With sh3 ' populate Match sheet by iterating through mentees
    For I = 1 To lngrows
    
        Call Match(1 + I, sh3.Cells(1 + I, 2).Value())
        
    Next I

End With

With sh5 ' populate Match sheet by iterating through mentees
    For I = 1 To lngrows
    
        Call Find_Mentor(sh5.Cells(1 + I, 6).Value(), 1 + I, lngrows)
        
    Next I

End With










End Sub


Public Function Create_Table()

Dim wb As Workbook
Dim sh1 As Worksheet
Dim sh2 As Worksheet
Dim sh3 As Worksheet
Dim sh4 As Worksheet
Dim sh5 As Worksheet
Dim sh6 As Worksheet




Set wb = ActiveWorkbook
Set sh1 = wb.Worksheets("Mentees")
Set sh2 = wb.Worksheets("Mentors")
Set sh3 = wb.Worksheets("Weight Matrix")
Set sh4 = wb.Worksheets("Category Weight Values")
Set sh5 = wb.Worksheets("Match")
Set sh6 = wb.Worksheets("mentors_used")

Dim cel As Range
Dim cel2 As Range
Dim cel3 As Range

Dim c1 As Integer
Dim c2 As Integer
Dim c3 As Integer

Dim SN1 As Variant
Dim SN2 As Variant
Dim SN3 As Variant

SN1 = "Student ID"                              'set SN1 to whatever you want to look for
SN2 = "I would be willing to mentor up to:"

With sh1.Range("A1:A100")           'set the range you want to look through
    Set cel = .Find(SN1, LookIn:=xlValues)
    'c1 = cel.column
    If cel3 Is Nothing Then
        c1 = 1
    Else
        c1 = cel.Column
    End If
End With


With sh2.Range("A1:A100")           'set the range you want to look through
    Set cel2 = .Find(SN1, LookIn:=xlValues)
    'c2 = cel2.column
    
    If cel3 Is Nothing Then
        Debug.Print "Name was not found."
        c2 = 1
    Else
        c2 = cel2.Column
    End If
    
    
    Set cel3 = .Find(SN2, LookIn:=xlValues)
    If cel3 Is Nothing Then
        c3 = 12
    Else
         c3 = cel3.Column
    End If
   ' c3 = cel3.column
    
End With

'Writing Mentee student number column in Weight matirx sheet and match sheet
Dim SrcRng1 As Range
With sh1
    Set SrcRng1 = .Range(.Cells(2, c1), .Cells(.Rows.Count, c1).End(xlUp))
    Set SrcRng4 = .Range(.Cells(2, c1 + 1), .Cells(.Rows.Count, c1 + 1).End(xlUp))
    Set SrcRng5 = .Range(.Cells(2, c1 + 2), .Cells(.Rows.Count, c1 + 2).End(xlUp))
    Set SrcRng6 = .Range(.Cells(2, c1 + 3), .Cells(.Rows.Count, c1 + 3).End(xlUp))
End With
sh3.Range("B2").Resize(SrcRng1.Rows.Count, 1).Value = SrcRng1.Value
'sh3.Range("B2:B300").SpecialCells(xlCellTypeBlanks).Delete
'sh3.Range("B1").Insert xlShiftDown

sh5.Range("A2").Resize(SrcRng1.Rows.Count, 1).Value = SrcRng1.Value
'sh5.Range("A2:A300").SpecialCells(xlCellTypeBlanks).Delete
'sh5.Range("A1").Insert xlShiftDown
'sh5.Cells(1, 1).Value() = "Mentee ID"

sh5.Range("B2").Resize(SrcRng4.Rows.Count, 1).Value = SrcRng4.Value
'sh5.Range("B2:B300").SpecialCells(xlCellTypeBlanks).Delete
'sh5.Range("B1").Insert xlShiftDown

sh5.Range("C2").Resize(SrcRng5.Rows.Count, 1).Value = SrcRng5.Value
'sh5.Range("C2:C300").SpecialCells(xlCellTypeBlanks).Delete
'sh5.Range("C1").Insert xlShiftDown

sh5.Range("D2").Resize(SrcRng6.Rows.Count, 1).Value = SrcRng6.Value
'sh5.Range("D2:D300").SpecialCells(xlCellTypeBlanks).Delete
'sh5.Range("D1").Insert xlShiftDown

'sh5.Cells(1, 1).Value() = "Mentee ID"
'sh5.Cells(1, 2).Value() = "Email"
'sh5.Cells(1, 3).Value() = "First Name"
'sh5.Cells(1, 4).Value() = "Last Name"



Dim lngRow As Long
Dim varArray As Variant

'Writing mentor student number column in weight matrix sheet
Dim SrcRng2 As Range
With sh2
    Set SrcRng2 = .Range(.Cells(2, c2), .Cells(.Rows.Count, c2).End(xlUp))
End With
varArray = SrcRng2.Value
With sh3
    For lngRow = 1 To UBound(varArray)
        sh3.Cells(1, lngRow + 2) = varArray(lngRow, 1)
    Next lngRow
        
End With
'sh3.Range("C1:C100").SpecialCells(xlCellTypeBlanks).Delete


'Writing Mentor student number column in match sheet
sh6.Range("A1").Resize(SrcRng2.Rows.Count, 1).Value = SrcRng2.Value
'sh6.Range("A1").SpecialCells(xlCellTypeBlanks).Delete

'Writing number of mentee for each mentor in match sheet
Dim SrcRng3 As Range
With sh2
    Set SrcRng3 = .Range(.Cells(2, c3), .Cells(.Rows.Count, c3).End(xlUp))
End With
sh6.Range("B1").Resize(SrcRng2.Rows.Count, 1).Value = SrcRng3.Value
'sh6.Range("B1").SpecialCells(xlCellTypeBlanks).Delete



'sh3.Range("B2").AutoFilter Field:=1, Criteria1:= _
       ' "<20000000", Operator:=xlOr, Criteria2:=">=30000000"
'sh3.Range("B2").SpecialCells(xlCellTypeVisible).Delete


End Function





Public Function Pair_Value(mentee As Long, mentor As Long, Lng As Integer, cell_row As Integer, cell_column As Integer)

Dim wb As Workbook
Dim sh1 As Worksheet
Dim sh2 As Worksheet
Dim sh3 As Worksheet
Dim sh4 As Worksheet
Dim sh5 As Worksheet
Dim sh6 As Worksheet




Set wb = ActiveWorkbook
Set sh1 = wb.Worksheets("Mentees")
Set sh2 = wb.Worksheets("Mentors")
Set sh3 = wb.Worksheets("Weight Matrix")
Set sh4 = wb.Worksheets("Category Weight Values")
Set sh5 = wb.Worksheets("Match")
Set sh6 = wb.Worksheets("mentors_used")
Dim match_value As Integer
match_value = 0


'find mentee in list
Dim cell As Range
Dim row_mentee As Long
Dim row_mentor As Long


With sh1.Columns("A:A")
    Set cell = .Find(mentee, LookIn:=xlValues)
End With

'If cell Is Nothing Then
'    'do it something
 '   MsgBox "Missing Student Number in Mentee List"
'    row_mentee = cell_row
'Else
'    row_mentee = cell.row
'End If
row_mentee = cell_row



With sh2.Columns("A:A")
    Set cell = .Find(mentor, LookIn:=xlValues)
End With

'If cell Is Nothing Then
 '   MsgBox "Missing Student Number in Mentor List"
 '   row_mentor = cell_column - 1
    
'Else
 '   row_mentor = cell.row
'End If
row_mentor = cell_column - 1

For I = 1 To Lng
      match_value = Similarity(sh1.Cells(row_mentee, I + 4).Value(), sh2.Cells(row_mentor, I + 4).Value()) * sh4.Cells(2, I).Value() + match_value
    Next I

sh3.Cells(cell_row, cell_column).Value() = match_value


End Function

Public Function Similarity(ByVal String1 As String, _
    ByVal String2 As String, _
    Optional ByRef RetMatch As String, _
    Optional min_match = 1) As Single
Dim b1() As Byte, b2() As Byte
Dim lngLen1 As Long, lngLen2 As Long
Dim lngResult As Long

If UCase(String1) = UCase(String2) Then
    Similarity = 1
Else:
    lngLen1 = Len(String1)
    lngLen2 = Len(String2)
    If (lngLen1 = 0) Or (lngLen2 = 0) Then
        Similarity = 0
    Else:
        b1() = StrConv(UCase(String1), vbFromUnicode)
        b2() = StrConv(UCase(String2), vbFromUnicode)
        lngResult = Similarity_sub(0, lngLen1 - 1, _
        0, lngLen2 - 1, _
        b1, b2, _
        String1, _
        RetMatch, _
        min_match)
        Erase b1
        Erase b2
        If lngLen1 >= lngLen2 Then
            Similarity = lngResult / lngLen1
        Else
            Similarity = lngResult / lngLen2
        End If
    End If
End If

End Function

Private Function Similarity_sub(ByVal start1 As Long, ByVal end1 As Long, _
                                ByVal start2 As Long, ByVal end2 As Long, _
                                ByRef b1() As Byte, ByRef b2() As Byte, _
                                ByVal FirstString As String, _
                                ByRef RetMatch As String, _
                                ByVal min_match As Long, _
                                Optional recur_level As Integer = 0) As Long
'* CALLED BY: Similarity *(RECURSIVE)

Dim lngCurr1 As Long, lngCurr2 As Long
Dim lngMatchAt1 As Long, lngMatchAt2 As Long
Dim I As Long
Dim lngLongestMatch As Long, lngLocalLongestMatch As Long
Dim strRetMatch1 As String, strRetMatch2 As String

If (start1 > end1) Or (start1 < 0) Or (end1 - start1 + 1 < min_match) _
Or (start2 > end2) Or (start2 < 0) Or (end2 - start2 + 1 < min_match) Then
    Exit Function '(exit if start/end is out of string, or length is too short)
End If

For lngCurr1 = start1 To end1
    For lngCurr2 = start2 To end2
        I = 0
        Do Until b1(lngCurr1 + I) <> b2(lngCurr2 + I)
            I = I + 1
            If I > lngLongestMatch Then
                lngMatchAt1 = lngCurr1
                lngMatchAt2 = lngCurr2
                lngLongestMatch = I
            End If
            If (lngCurr1 + I) > end1 Or (lngCurr2 + I) > end2 Then Exit Do
        Loop
    Next lngCurr2
Next lngCurr1

If lngLongestMatch < min_match Then Exit Function

lngLocalLongestMatch = lngLongestMatch
RetMatch = ""

lngLongestMatch = lngLongestMatch _
+ Similarity_sub(start1, lngMatchAt1 - 1, _
start2, lngMatchAt2 - 1, _
b1, b2, _
FirstString, _
strRetMatch1, _
min_match, _
recur_level + 1)
If strRetMatch1 <> "" Then
    RetMatch = RetMatch & strRetMatch1 & "*"
Else
    RetMatch = RetMatch & IIf(recur_level = 0 _
    And lngLocalLongestMatch > 0 _
    And (lngMatchAt1 > 1 Or lngMatchAt2 > 1) _
    , "*", "")
End If


RetMatch = RetMatch & Mid$(FirstString, lngMatchAt1 + 1, lngLocalLongestMatch)


lngLongestMatch = lngLongestMatch _
+ Similarity_sub(lngMatchAt1 + lngLocalLongestMatch, end1, _
lngMatchAt2 + lngLocalLongestMatch, end2, _
b1, b2, _
FirstString, _
strRetMatch2, _
min_match, _
recur_level + 1)

If strRetMatch2 <> "" Then
    RetMatch = RetMatch & "*" & strRetMatch2
Else
    RetMatch = RetMatch & IIf(recur_level = 0 _
    And lngLocalLongestMatch > 0 _
    And ((lngMatchAt1 + lngLocalLongestMatch < end1) _
    Or (lngMatchAt2 + lngLocalLongestMatch < end2)) _
    , "*", "")
End If

Similarity_sub = lngLongestMatch

End Function

Public Function Match(cell_row As Integer, mentee As Long)


Dim wb As Workbook
Dim sh1 As Worksheet
Dim sh2 As Worksheet
Dim sh3 As Worksheet
Dim sh4 As Worksheet
Dim sh5 As Worksheet
Dim sh6 As Worksheet




Set wb = ActiveWorkbook
Set sh1 = wb.Worksheets("Mentees")
Set sh2 = wb.Worksheets("Mentors")
Set sh3 = wb.Worksheets("Weight Matrix")
Set sh4 = wb.Worksheets("Category Weight Values")
Set sh5 = wb.Worksheets("Match")
Set sh6 = wb.Worksheets("mentors_used")


Dim mentor As Long
mentor = 0

Dim M_left As Long ' number of mentees a mentor can still be assigned
M_left = 0

Dim match_score As Integer
match_score = 0

Dim c As Long 'column
Dim r As Long 'row

Dim c_mentor As Long 'mentor column
c_mentor = cell_column + 1



With sh3
    lngcolumn = sh3.Cells(1, Columns.Count).End(xlToLeft).Column - 2
    For I = 1 To lngcolumn
        If max_score <= sh3.Cells(cell_row, I + 2).Value() Then
            max_score = sh3.Cells(cell_row, I + 2).Value()
            c_mentor = I + 2
        End If
    Next I
    
End With

mentor = sh3.Cells(1, c_mentor).Value()
With sh6.Range("A1:E65536")           'set the range you want to look through
    Set cel = .Find(mentor, LookIn:=xlValues)
End With

If cel Is Nothing Then
    'do it something
   ' MsgBox "Missing Student Number in Mentor List in mentors_used sheet"
    
Else
    c = cel.Column
    r = cel.Row
    M_left = sh6.Cells(r, 2).Value()
    M_left = M_left - 1
    sh6.Cells(r, 2).Value = M_left

    If M_left < 1 Then
        sh6.Cells(r, 1).Delete
        sh6.Cells(r, 2).Delete
        sh3.Columns(c_mentor).EntireColumn.Delete
    End If

End If
   



'With sh5.Range("A1:E65536")           'set the range you want to look through
 '   Set cel = .Find(mentee, LookIn:=xlValues)
'End With

'If cel Is Nothing Then
    'do it something
  '  c = 1
 '   r = cell_row - 1
'Else
 '   c = cel.Column
  '  r = cel.row
'End If
r = cell_row

sh5.Cells(r, 6).Value() = mentor ' write Mentor for mentee in table in Match sheet



End Function

Public Function Find_Mentor(mentor As Long, M_row As Integer, Lng As Integer)


Dim wb As Workbook
Dim sh1 As Worksheet
Dim sh2 As Worksheet
Dim sh3 As Worksheet
Dim sh4 As Worksheet
Dim sh5 As Worksheet
Dim sh6 As Worksheet




Set wb = ActiveWorkbook
Set sh1 = wb.Worksheets("Mentees")
Set sh2 = wb.Worksheets("Mentors")
Set sh3 = wb.Worksheets("Weight Matrix")
Set sh4 = wb.Worksheets("Category Weight Values")
Set sh5 = wb.Worksheets("Match")
Set sh6 = wb.Worksheets("mentors_used")

Dim r As Integer
r = 0


For I = 1 To Lng
    If sh2.Cells(I, 1).Value() = mentor Then
        r = I
    End If
Next I

If r = 0 Then
    'Do something
Else
    sh5.Cells(M_row, 7).Value() = sh2.Cells(r, 2).Value()
    sh5.Cells(M_row, 8).Value() = sh2.Cells(r, 3).Value()
    sh5.Cells(M_row, 9).Value() = sh2.Cells(r, 4).Value()
    
End If

    








End Function


