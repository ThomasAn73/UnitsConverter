Attribute VB_Name = "FileIO"
Option Explicit

Public Const Delimiter = "¶"
Dim UnitsIniFile As String

Public Sub LoadUnitsFromFile()
    'On Error GoTo FileError
    On Error Resume Next
    Err.Clear
    
    Dim FileNum As Integer
    Dim OneFileLine As String
    Dim Fragment
    Dim Counter As Integer
    Dim UnitsUbound As Integer 'this starts from 0
    Dim CurrentSection As Integer
    Dim Integrity As Boolean
    Dim EndOfFile As Boolean
    
    FileNum = FreeFile
    UnitsIniFile = App.Path & "\Units.ini"
    Open UnitsIniFile For Input As FileNum
    
    CurrentSection = 0
    UnitsUbound = -1
    CategoriesUbound = -1
    EndOfFile = False
    
    'Retrieve categories
    Do While (Err.Number = 0 And EndOfFile = False)
        Line Input #FileNum, OneFileLine
        OneFileLine = Trim(OneFileLine)
        EndOfFile = EOF(FileNum)
        If (Len(OneFileLine) > 0 And Left(OneFileLine, 1) <> "#") Then 'Skip comment lines
            Fragment = Split(OneFileLine, Delimiter)
            If (LCase(Trim(Fragment(0))) = "section=categories") Then 'Entering Categories section
                CurrentSection = CurrentSection + 1
                CategoriesUbound = -1
            ElseIf (LCase(Trim(Fragment(0))) = "section=units") Then 'Entering Units Section
                CurrentSection = CurrentSection + 1
                If (CurrentSection = 1) Then
                    Err.Raise -1, , "Units.ini file is corrupt." & vbCr & "Categories section not on top."
                    Exit Do
                End If
                UnitsUbound = -1
            ElseIf (CurrentSection = 1 And UBound(Fragment) >= 1) Then 'Retrieve Categories
                CategoriesUbound = CategoriesUbound + 1
                ReDim Preserve UnitsDataBase(CategoriesUbound)
                UnitsDataBase(CategoriesUbound).Category.CategoryID = Trim(Fragment(0))
                UnitsDataBase(CategoriesUbound).Category.CategoryName = Trim(Fragment(1))
                UnitsDataBase(CategoriesUbound).Category.RelatedUnitCount = 0
            ElseIf (CurrentSection = 2 And UBound(Fragment) >= 4) Then 'Retrieve Units
                For Counter = 0 To UBound(UnitsDataBase)
                    If (UnitsDataBase(Counter).Category.CategoryID = Trim(Fragment(1))) Then
                        UnitsDataBase(Counter).Category.RelatedUnitCount = UnitsDataBase(Counter).Category.RelatedUnitCount + 1
                        UnitsUbound = UnitsDataBase(Counter).Category.RelatedUnitCount - 1
                        
                        ReDim Preserve UnitsDataBase(Counter).RelatedUnits(UnitsUbound)
                        UnitsDataBase(Counter).RelatedUnits(UnitsUbound).Serial = Trim(Fragment(0))
                        UnitsDataBase(Counter).RelatedUnits(UnitsUbound).LinkedToCat = Trim(Fragment(1))
                        UnitsDataBase(Counter).RelatedUnits(UnitsUbound).LongName = Trim(Fragment(2))
                        UnitsDataBase(Counter).RelatedUnits(UnitsUbound).ShortName = Trim(Fragment(3))
                        UnitsDataBase(Counter).RelatedUnits(UnitsUbound).ConversionFactor = Replace(Fragment(4), " ", "") 'remove *all* spaces (not just end spaces)
                        If (UBound(Fragment) >= 5) Then UnitsDataBase(Counter).RelatedUnits(UnitsUbound).Offset = Replace(Fragment(5), " ", "")
                        If (UBound(Fragment) >= 6) Then UnitsDataBase(Counter).RelatedUnits(UnitsUbound).UnitSystem = Trim(Fragment(6))
                        If (UBound(Fragment) >= 7) Then UnitsDataBase(Counter).RelatedUnits(UnitsUbound).Description = Trim(Fragment(7))
                        Exit For
                    End If
                Next
            End If
        End If
    Loop
    Close #FileNum
    
    'Do something if the ini file contained no data
    '.....No leave it alone
    
    'Do something if there was no ini file at all
    If (Err.Number = 53) Then Call SaveUnitsToFile

    'Issue some message
    If (Err.Number <> 0) Then MsgBox Err.Description & vbCr & "ErrCode: " & Err.Number, vbOKOnly, "Error"
End Sub

Public Sub SaveUnitsToFile()
    Dim OneFileLine As String
    Dim FileNum As Integer
    Dim Count As Integer
    Dim count2 As Integer
    
    FileNum = FreeFile
    Open UnitsIniFile For Output As FileNum
    
    'Categories section
    OneFileLine = "Section=Categories" & vbCrLf & "#Sectiom Format-> CustomSerial¶ CatName"
    Print #FileNum, OneFileLine
    For Count = 0 To CategoriesUbound
        OneFileLine = UnitsDataBase(Count).Category.CategoryID & "¶ " & UnitsDataBase(Count).Category.CategoryName
        Print #FileNum, OneFileLine
    Next
    OneFileLine = ""
    Print #FileNum, OneFileLine
    
    'Units Section
    OneFileLine = "Section=Units" & vbCrLf & "#Section Format-> CustomSerial¶ LinkToCategorySerial¶ UnitName¶ UnitAbbreviation¶ ConversionFactor¶ Offset¶ UnitSystem¶ description"
    Print #FileNum, OneFileLine
    For Count = 0 To CategoriesUbound 'categories
        For count2 = 0 To UnitsDataBase(Count).Category.RelatedUnitCount - 1 'units per category
            OneFileLine = UnitsDataBase(Count).RelatedUnits(count2).Serial & "¶ " & _
            UnitsDataBase(Count).RelatedUnits(count2).LinkedToCat & "¶ " & _
            UnitsDataBase(Count).RelatedUnits(count2).LongName & "¶ " & _
            UnitsDataBase(Count).RelatedUnits(count2).ShortName & "¶ " & _
            UnitsDataBase(Count).RelatedUnits(count2).ConversionFactor & "¶ " & _
            UnitsDataBase(Count).RelatedUnits(count2).Offset & "¶ " & _
            UnitsDataBase(Count).RelatedUnits(count2).UnitSystem & "¶ " & _
            UnitsDataBase(Count).RelatedUnits(count2).Description
            Print #FileNum, OneFileLine
        Next
        OneFileLine = ""
        Print #FileNum, OneFileLine
    Next
    
    Close #FileNum
End Sub
