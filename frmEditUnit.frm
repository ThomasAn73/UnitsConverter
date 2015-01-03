VERSION 5.00
Begin VB.Form frmEditUnit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Unit"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox NewEdit 
      Height          =   285
      Index           =   5
      Left            =   3000
      TabIndex        =   15
      Text            =   "Text1"
      ToolTipText     =   "This is mostly used for temperature conversions."
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   1
      Left            =   4785
      TabIndex        =   11
      Top             =   45
      Width           =   650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   10
      Top             =   45
      Width           =   650
   End
   Begin VB.TextBox NewEdit 
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      ToolTipText     =   "The description will be seen as a tooltip (for this unit) in the list."
      Top             =   1380
      Width           =   5295
   End
   Begin VB.TextBox NewEdit 
      Height          =   285
      Index           =   3
      Left            =   4560
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox NewEdit 
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text1"
      ToolTipText     =   "Number, or expression."
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox NewEdit 
      Height          =   285
      Index           =   1
      Left            =   3960
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox NewEdit 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   360
      Index           =   5
      Left            =   3045
      TabIndex        =   17
      Top             =   975
      Width           =   1770
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "+Offset"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   16
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   360
      Index           =   2
      Left            =   1470
      TabIndex        =   14
      Top             =   975
      Width           =   1500
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   1
      Left            =   1020
      TabIndex        =   13
      Top             =   60
      Width           =   2925
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   120
      X2              =   5400
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label2 
      Caption         =   "Parent Unit:"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   0
      Left            =   135
      TabIndex        =   12
      Top             =   60
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Description"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   1140
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "UnitSystem"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ConversionFactor"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Symbol"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmEditUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Deactivate()
    frmEditUnit.SetFocus
End Sub

Private Sub Form_Load()
    EditMode = True
    frmMainWindow.Enabled = False
    
    If (frmMainWindow.Check2 = 1) Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3 'On top of all windows Else SetWindowPos Me.hWnd, -2, 0, 0, 0, 0, 3
    
    'Position the window (center)
    Me.Left = frmMainWindow.Left + (frmMainWindow.Width - Me.Width) / 2
    Me.Top = frmMainWindow.Top + (frmMainWindow.Height - Me.Height) / 2
    
    'Populate the Text fields with the active timer data
    NewEdit(0).Text = UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Selection.UnitDatabaseIndex).LongName
    NewEdit(1).Text = UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Selection.UnitDatabaseIndex).ShortName
    NewEdit(2).Text = UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Selection.UnitDatabaseIndex).ConversionFactor
    NewEdit(3).Text = UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Selection.UnitDatabaseIndex).UnitSystem
    NewEdit(4).Text = UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Selection.UnitDatabaseIndex).Description
    NewEdit(5).Text = UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Selection.UnitDatabaseIndex).Offset
    If (ParentUnit.Count = 1) Then
        Label2(1).Caption = UnitsDataBase(ParentUnit.CategoryDatabaseIndex).RelatedUnits(ParentUnit.UnitDatabaseIndex).LongName
        Label1(2).Caption = "1 " & UnitsDataBase(ParentUnit.CategoryDatabaseIndex).RelatedUnits(ParentUnit.UnitDatabaseIndex).LongName & " ="
    ElseIf (ParentUnit.Count > 1) Then
        'Label2(1).Caption = "Unknown (multiple found)"
        'Label1(2).Caption = "1 Unknown unit ="
        Label2(1).Caption = ParentUnit.Names
        Label1(2).Caption = "1 " & Trim(Split(ParentUnit.Names, ",")(0)) & " ="
    Else
        Label2(1).Caption = "Unknown"
        Label1(2).Caption = "1 Unknown unit ="
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim Expression As Variant
   
    Select Case Index
        Case 1 'cancel
            Unload Me
        Case 0 'apply
        
            'Evaluate the conversion box expression
            Expression = EvalExpression(NewEdit(2).Text)
            If (Expression(1) <> 0) Then 'if there is an error code
                Label3(2).Caption = Expression(2)
                Exit Sub
            ElseIf (Expression(0) = 0) Then ' Should not be zero
                Label3(2).Caption = "Result cannot be zero"
                Exit Sub
            End If
            
            Expression = EvalExpression(NewEdit(5).Text)
            If (Expression(1) <> 0) Then 'if there is an error code
                Label3(5).Caption = Expression(2)
                Exit Sub
            End If
                     
            'Copy the unit's edited data into the database
            UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Selection.UnitDatabaseIndex).LongName = Trim(NewEdit(0).Text)
            UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Selection.UnitDatabaseIndex).ShortName = Trim(NewEdit(1).Text)
            UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Selection.UnitDatabaseIndex).ConversionFactor = NewEdit(2).Text
            UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Selection.UnitDatabaseIndex).UnitSystem = Trim(NewEdit(3).Text)
            UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Selection.UnitDatabaseIndex).Description = Trim(NewEdit(4).Text)
            UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Selection.UnitDatabaseIndex).Offset = NewEdit(5).Text
            Call SaveUnitsToFile
            
            'Update the corresponding listview item
            frmMainWindow.ListView1(1).SelectedItem.Text = Trim(NewEdit(0).Text)
            frmMainWindow.ListView1(1).SelectedItem.ListSubItems(2).Text = Trim(NewEdit(1).Text)
            frmMainWindow.ListView1(1).SelectedItem.ListSubItems(3).Text = Trim(NewEdit(3).Text)
            frmMainWindow.ListView1(1).SelectedItem.ToolTipText = Trim(NewEdit(4).Text)
            
            Call frmMainWindow.SortList(1)
            Call frmMainWindow.RecalcValues
            Unload Me
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EditMode = False
    frmMainWindow.Enabled = True
End Sub

Private Sub NewEdit_Change(Index As Integer)
    Dim ExpressionResult
    Select Case Index
        Case 2, 5
            ExpressionResult = EvalExpression(NewEdit(Index).Text)
            If (ExpressionResult(1) = 0) Then
                Label3(Index).ForeColor = &H80000011
                If (ExpressionResult(3) > 1) Then Label3(Index).Caption = "=" & Round(ExpressionResult(0), 14) Else Label3(Index).Caption = ""
            Else
                Label3(Index).ForeColor = RGB(220, 0, 0)
                Label3(Index).Caption = ExpressionResult(2)
            End If
    End Select
End Sub

'-----------------------------------------------------------------------
'OBSOLETE CODE section--------------------------------------------------
'-----------------------------------------------------------------------

'These proceedures were being used for dynamic key guarding during expression entry

Private Sub NewEdit_KeyDownOLD(Index As Integer, KeyCode As Integer, Shift As Integer)
    TrackExpSyntax.LastKeyCode = KeyCode
    If (KeyCode = 46) Then Call NewEdit_KeyPressOLD(Index, 0 - KeyCode)
End Sub

Private Sub NewEdit_lostFocusOLD(Index As Integer)
    Select Case Index
        Case 2
            Call NewEdit_KeyPressOLD(2, 0) 'This will force a new parenthesis count
        Case 5
            Call NewEdit_KeyPressOLD(5, 0) 'This will force a new parenthesis count
    End Select
End Sub

Private Sub NewEdit_KeyPressOLD(Index As Integer, KeyAscii As Integer)
    Dim CharCheck As Variant
    Dim Count As Integer
    Dim ClipboardStr As String
    
    If (KeyAscii = 13) Then 'accept values by pressing enter
        '.... also do a final check for the validity of the expression in the conversion box
        Command1_Click (0)
        Exit Sub
    ElseIf (KeyAscii = 3 Or KeyAscii = 26) Then 'user pressed CNTR+C, or CTRL+Z ... let it go through
        Exit Sub
    End If
    
    Select Case Index
        Case 0 ' name box
        Case 1 ' symbol box
        Case 2, 5 ' conversion factor box, or offset Box
        
            'do various dynamic checks for allowable char entry
            If (KeyAscii = 22) Then 'User pressed CNTR+V ... try to paste the clipboard string one char at the time
                ClipboardStr = Trim(Clipboard.GetText)
                If (ClipboardStr = "") Then Exit Sub
                For Count = 1 To Len(ClipboardStr)
                    CharCheck = AsciiAfterExpressionCheck(Asc(Mid(ClipboardStr, Count, 1)), NewEdit(Index).Text & Left(ClipboardStr, Count - 1), NewEdit(Index).SelStart + Count - 1, NewEdit(Index).SelLength - (Count - 1))
                    If (CharCheck(0) = 0) Then
                        KeyAscii = 0
                        Exit For
                    End If
                Next
            Else
                CharCheck = AsciiAfterExpressionCheck(KeyAscii, NewEdit(Index).Text, NewEdit(Index).SelStart, NewEdit(Index).SelLength)
                KeyAscii = CharCheck(0)
            End If
            Label3(Index).ForeColor = RGB(220, 0, 0)
            If (CharCheck(1) = "") Then Label3(Index).Caption = CharCheck(2) Else Label3(Index).Caption = CharCheck(1)
        
        Case 3 ' UnitSystem box
        Case 4 ' Description box
            
    End Select
    
End Sub
