VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMainWindow 
   Caption         =   "Units Conversion Tool"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7380
   Icon            =   "frmMainWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Pin"
      Height          =   255
      Left            =   195
      TabIndex        =   14
      ToolTipText     =   "Pin Window on Top of all others on the desktop"
      Top             =   105
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Decimals"
      Height          =   255
      Left            =   5610
      TabIndex        =   13
      Top             =   75
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   6615
      TabIndex        =   7
      Text            =   "4"
      Top             =   60
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   3285
      TabIndex        =   5
      Text            =   "1"
      ToolTipText     =   "Number, or expression for selected unit. Hover mouse over the other units and CTRL+C to copy their corresponding value."
      Top             =   60
      Width           =   1950
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3720
      Index           =   1
      Left            =   1920
      TabIndex        =   12
      Top             =   735
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6562
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label4 
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   3450
      TabIndex        =   15
      Top             =   480
      Width           =   3270
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   6780
      TabIndex        =   11
      Top             =   510
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   7005
      TabIndex        =   10
      Top             =   510
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   1245
      TabIndex        =   9
      Top             =   510
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1470
      TabIndex        =   8
      Top             =   510
      Width           =   195
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   180
      X2              =   7200
      Y1              =   390
      Y2              =   390
   End
   Begin VB.Label Label1 
      Caption         =   "Value of selected unit"
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   6
      Top             =   90
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Related Units"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   1
      Left            =   1995
      TabIndex        =   4
      Top             =   480
      Width           =   1290
   End
   Begin VB.Label Label2 
      Caption         =   "Version:"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   1
      Left            =   165
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   2
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Category"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Menu ContextMenu 
      Caption         =   "ContextMenu"
      Visible         =   0   'False
      Begin VB.Menu ContextMenuItem 
         Caption         =   "Add"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum Menu
    Add = 0
    Delete = 1
    Edit = 2
End Enum

Enum Color
    black = &H80000012
    blue = &HFF0000
    grey = &H80000011
    darkgrey = &H80000015
    Highlight = &H8000000D
End Enum

Dim ActiveListView As Integer
Dim IsWindowActive As Boolean

Private Sub Form_Load()
    Label2(1).Caption = "Version: 20080525"
    Label2(0).Caption = "Thomas Anagnostou - Rayflectar Graphics"
    
    Text1(0).Tag = Val(Text1(0).Text) & Delimiter
    
    Call InitializeContextMenu
    Call FileIO.LoadUnitsFromFile
    Call PopulateViews
    
    Call Check1_Click
    EditMode = False
End Sub

Private Sub Form_Resize()
    If (frmMainWindow.WindowState <> 0) Then Exit Sub 'To prevent Vb error. Somehow the resize event is triggered when minimizing the window and causes a VB error.
    
    If (frmMainWindow.Width < 7500) Then
        frmMainWindow.Width = 7500
        frmMainWindow.Enabled = False
    ElseIf (frmMainWindow.Height < 5340) Then
        frmMainWindow.Height = 5340
        frmMainWindow.Enabled = False
    End If
    
    frmMainWindow.Enabled = True
    Call AdjustWindow
End Sub

'-----------------------------------------------------------------------
'MOUSE events ----------------------------------------------------------
'-----------------------------------------------------------------------

Private Sub Label3_Click(Index As Integer)
    Select Case Index
        Case 0 'add (from listview1(0) )
            ListView1(0).SetFocus
            ListView1_GotFocus (0)
            Call ContextMenuItem_Click(Menu.Add)
        Case 1 'delete (from listview1(0) )
            ListView1(0).SetFocus
            ListView1_GotFocus (0)
            Call ContextMenuItem_Click(Menu.Delete)
        Case 2 'add (from listview1(1) )
            ListView1(1).SetFocus
            ListView1_GotFocus (1)
            Call ContextMenuItem_Click(Menu.Add)
        Case 3 'add (from listview1(1) )
            ListView1(1).SetFocus
            ListView1_GotFocus (1)
            Call ContextMenuItem_Click(Menu.Delete)
    End Select
End Sub

Private Sub Check1_Click()
    Text1(1).Enabled = Check1.Value
    Call Text1_Change(1)
End Sub

Private Sub Check2_Click()
    If (Check2.Value = 1) Then
        SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3 'On top of all windows
    Else
        SetWindowPos Me.hWnd, -2, 0, 0, 0, 0, 3 'On top of all windows
    End If
End Sub



'Handles changes in highliting (trigers recalculation of unit values for display)
Private Sub ListView1_Click(Index As Integer)
    Select Case Index
        Case 0 'User Clicked in Categories listview
            Call SortList(1) 'sort the other panel
        Case 1 'User Clicked in Units Listview
            Call SortList(0) 'sort the other panel
    End Select
End Sub


Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Select Case Index
        Case 0
            Call ShowRelatedUnits
        Case 1
            Call RecalcValues
    End Select
End Sub

'Pops up the context menu
Private Sub listview1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Check for RightMouseClick
    If (Button <> 2) Then Exit Sub
    If (ListView1(Index).ListItems.Count < 1) Then
        ContextMenuItem(Menu.Delete).Enabled = False
        ContextMenuItem(Menu.Edit).Enabled = False
    Else
        ContextMenuItem(Menu.Delete).Enabled = True
        ContextMenuItem(Menu.Edit).Enabled = True
    End If
    PopupMenu ContextMenu
End Sub

Private Sub ListView1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim MousedItem As Object
    Dim SomeString As String
    
    If (GetForegroundWindow <> Me.hWnd) Then Exit Sub
    Select Case Index
        Case 1 'Units listview
            If (Not (ListView1(1).HitTest(X, Y) Is Nothing)) Then
                If (Val(Label4.Tag) <> ListView1(1).HitTest(X, Y).Index) Then
                    ListView1(1).SetFocus
                    Call FeedbackString(False, ListView1(1).HitTest(X, Y).ListSubItems(1).Text & " " & ListView1(1).HitTest(X, Y).ListSubItems(2).Text, ListView1(1).HitTest(X, Y).Index)
                End If
                
                If (X > ListView1(1).ColumnHeaders(1).Width And ListView1(1).ToolTipText <> ListView1(1).HitTest(X, Y).ToolTipText) Then
                    ListView1(1).ToolTipText = ListView1(1).HitTest(X, Y).ToolTipText
                ElseIf (X <= ListView1(1).ColumnHeaders(1).Width And ListView1(1).ToolTipText <> "") Then
                    ListView1(1).ToolTipText = ""
                End If
            ElseIf (Label4.Tag <> "") Then
                Call FeedbackString
            End If
    End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SomeString As String
    'If the form detects movement it means you are no longer inside the listview boundaries (so we need to clear the feedback message
    If (Label4.Tag <> "") Then 'The label4.tag usually contains the index of the listitem that the mouse is hovering on
        Call FeedbackString
    End If
End Sub

'-----------------------------------------------------------------------
'FOCUS events ----------------------------------------------------------
'-----------------------------------------------------------------------

Private Sub ListView1_GotFocus(Index As Integer)
    ActiveListView = Index
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Dim SomeString As String

    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    Select Case Index
        Case 0
            If (UBound(Split(Text1(0).Tag, Delimiter)) > 0) Then SomeString = Split(Text1(0).Tag, Delimiter)(1) Else SomeString = ""
            Call FeedbackString(True, SomeString)
    End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Call FeedbackString(False)
End Sub

'-----------------------------------------------------------------------
'CHANGE events ---------------------------------------------------------
'-----------------------------------------------------------------------

Private Sub Text1_Change(Index As Integer)
    Dim ExpressionResult As Variant
    
    Select Case Index
        Case 0 ' Active value
            
            'Evaluate the expression in the textbox
            ExpressionResult = EvalExpression(Text1(Index).Text)
            If (ExpressionResult(1) = 0) Then
                Text1(0).Tag = ExpressionResult(0) & Delimiter
                Call FeedbackString(True)
            Else
                Text1(0).Tag = Split(Text1(0).Tag, Delimiter)(0) & Delimiter & ExpressionResult(2)
                Call FeedbackString(True, CStr(ExpressionResult(2)))
            End If
        Case 1 ' Decimals
            Text1(Index).Text = Int(Val(Text1(Index).Text))
            If (Val(Text1(Index)) < 0) Then Text1(Index).Text = 0
            If (Val(Text1(Index)) > 20) Then Text1(Index).Text = 20
            If (Check1.Value = 0) Then
                DecimalRounding = "General Number" '"0.000 E+00"
            ElseIf (Val(Text1(1).Text) = 0) Then
                DecimalRounding = "0"
            Else
                DecimalRounding = "0." & String(Val(Text1(1).Text), "0")
            End If
    End Select
    Call RecalcValues
End Sub

'-----------------------------------------------------------------------
'KEY PRESS events ------------------------------------------------------
'-----------------------------------------------------------------------

Private Sub ListView1_Keydown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 113, 13 'Pressing the F2 key or enter
            If (Not (ListView1(Index).SelectedItem Is Nothing)) Then ListView1(Index).StartLabelEdit
        Case 46
            Call ContextMenuItem_Click(Menu.Delete)
        Case 38, 40 '(up and down arrows... now handled by the itemclick event)
    End Select
End Sub

Private Sub ListView1_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 1 'Units listview
            If (KeyAscii = 3) Then
                Clipboard.Clear
                Clipboard.SetText Trim(Label4.Caption)
                Label4.Alignment = 0
                If (Label4.Tag <> "") Then Label4.Caption = ListView1(1).ListItems(Val(Label4.Tag)) & "->copied to clipboard"
            End If
    End Select
End Sub

'-----------------------------------------------------------------------
'LABEL EDIT events -----------------------------------------------------
'-----------------------------------------------------------------------

'Handles trivial label edits for list items (and updates the database)
Private Sub ListView1_AfterLabelEdit(Index As Integer, Cancel As Integer, NewString As String)
    If (Cancel = 1 Or NewString = ListView1(Index).SelectedItem.Text) Then Exit Sub
    If (Index = 0) Then
        UnitsDataBase(Val(ListView1(0).SelectedItem.Key)).Category.CategoryName = NewString
    ElseIf (Index = 1) Then
        UnitsDataBase(Val(ListView1(0).SelectedItem.Key)).RelatedUnits(Val(ListView1(1).SelectedItem.Key)).LongName = NewString
    End If
    ListView1(Index).SelectedItem.Text = NewString
    Call ShowRelatedUnits 'This is to update the Listview1(1) label text (the sub will exit immediately after that)
    Call SaveUnitsToFile
End Sub

'-----------------------------------------------------------------------
'CONTEXT MENU procedures / events --------------------------------------
'-----------------------------------------------------------------------

Private Sub InitializeContextMenu()
    'Populate context menu with vaious commands
    'Context menu is a control, but you cannot add it via drag/drop into the form. You have to use the menu editor
    'By use of the menu editor the first menu item is already inserted. The rest can be added programatically (as for example in here below)
    
    'Load ContextMenuItem(Menu.Add) 'This object is already loaded by the menu editor
    ContextMenuItem(Menu.Add).Visible = True
    ContextMenuItem(Menu.Add).Enabled = True
    'ContextMenuItem(Menu.Add).Caption = "Add"
    
    Load ContextMenuItem(Menu.Delete)
    ContextMenuItem(Menu.Delete).Visible = True
    ContextMenuItem(Menu.Delete).Enabled = True
    ContextMenuItem(Menu.Delete).Caption = "Delete"
    
    Load ContextMenuItem(Menu.Edit)
    ContextMenuItem(Menu.Edit).Visible = True
    ContextMenuItem(Menu.Edit).Enabled = True
    ContextMenuItem(Menu.Edit).Caption = "Edit"
End Sub

'Decides what to do when a context menu selection is made
Private Sub ContextMenuItem_Click(Index As Integer)
    Dim Count As Integer
    Dim count2 As Integer
    Dim count3 As Long
    Dim UserDecision
    Dim Message As String
    
    Select Case Index
        Case Menu.Add
        '-----------------------------------------------
            If (ActiveListView = 0) Then 'Add a category
                'adjust the counter
                CategoriesUbound = CategoriesUbound + 1
                
                'Expand the Database
                ReDim Preserve UnitsDataBase(CategoriesUbound)
                UnitsDataBase(CategoriesUbound).Category.RelatedUnitCount = 0
                UnitsDataBase(CategoriesUbound).Category.CategoryName = "Untitled"
                
                'Assign a category serial (find the largest existing serial and add 10 to get a new one)
                For Count = 0 To CategoriesUbound
                    If (UnitsDataBase(Count).Category.CategoryID > count2) Then count2 = UnitsDataBase(Count).Category.CategoryID
                Next
                UnitsDataBase(CategoriesUbound).Category.CategoryID = count2 + 10
                
                'Save the database
                Call SaveUnitsToFile
                
                'Add item in the listview
                ListView1(0).Sorted = False
                ListView1(0).ListItems.Add , CategoriesUbound & "ID", UnitsDataBase(CategoriesUbound).Category.CategoryName
            
                'Update the selection variables
                ListView1(0).ListItems(CategoriesUbound + 1).Selected = True
                Selection.CategoryDatabaseIndex = CategoriesUbound
                Selection.CategoryListviewRow = CategoriesUbound + 1
                Selection.CategoryUnitsCount = 0
                Selection.Exists = True
                Selection.UnitDatabaseIndex = -1
                Selection.UnitListviewRow = 0
                
                Call SortList(0)
                Call ShowRelatedUnits(True)
        '-----------------------------------------------
            ElseIf (ActiveListView = 1) Then 'Add a unit
                'adjust the counter in the database itself
                UnitsDataBase(Selection.CategoryDatabaseIndex).Category.RelatedUnitCount = ListView1(1).ListItems.Count + 1
            
                'Expand the Database
                ReDim Preserve UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(UnitsDataBase(Selection.CategoryDatabaseIndex).Category.RelatedUnitCount - 1)
                UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(UnitsDataBase(Selection.CategoryDatabaseIndex).Category.RelatedUnitCount - 1).LongName = "Untitled"
                UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(UnitsDataBase(Selection.CategoryDatabaseIndex).Category.RelatedUnitCount - 1).ShortName = ""
                UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(UnitsDataBase(Selection.CategoryDatabaseIndex).Category.RelatedUnitCount - 1).LinkedToCat = UnitsDataBase(Selection.CategoryDatabaseIndex).Category.CategoryID
                UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(UnitsDataBase(Selection.CategoryDatabaseIndex).Category.RelatedUnitCount - 1).ConversionFactor = 10
                UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(UnitsDataBase(Selection.CategoryDatabaseIndex).Category.RelatedUnitCount - 1).Offset = 0
                UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(UnitsDataBase(Selection.CategoryDatabaseIndex).Category.RelatedUnitCount - 1).Description = ""
                UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(UnitsDataBase(Selection.CategoryDatabaseIndex).Category.RelatedUnitCount - 1).UnitSystem = ""
                
                'Assign a unit serial (find the largest existing serial and add 10 to get a new one)
                For Count = 0 To CategoriesUbound 'step through each category
                    For count2 = 0 To UnitsDataBase(Count).Category.RelatedUnitCount - 1
                        If (UnitsDataBase(Count).RelatedUnits(count2).Serial > count3) Then count3 = UnitsDataBase(Count).RelatedUnits(count2).Serial
                    Next
                Next
                UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(UnitsDataBase(Selection.CategoryDatabaseIndex).Category.RelatedUnitCount - 1).Serial = count3 + 10
                
                'Save the database
                Call SaveUnitsToFile
                
                Call ShowRelatedUnits(True, ListView1(1).ListItems.Count + 1)
                'Call EditUnit
            End If
        
        Case Menu.Delete
        '-----------------------------------------------
            If (ListView1(ActiveListView).ListItems.Count = 0) Then Exit Sub
            If (ActiveListView = 0) Then 'Delete whole category
                Message = "Delete the selected Category <" & ListView1(0).SelectedItem.Text & "> and all associated units ?"
                UserDecision = MsgBox(Message, vbOKCancel + vbDefaultButton2, "Confirm")
                If (UserDecision <> 1) Then Exit Sub
                
                'adjust the counter
                CategoriesUbound = CategoriesUbound - 1
                
                'Remove the item from the listview (it is important to do it up here, otherwise you will have problems with key assignments)
                ListView1(0).ListItems.Remove (Selection.CategoryListviewRow) 'This will cause automatic re-indexing of the listview
                
                'Shift the database down by one
                'if the deleted item was last on a long list then this for-loop is being bypassed
                For Count = Selection.CategoryDatabaseIndex To CategoriesUbound 'go up to the previous to last item (the categoriesUbound is new so it is already one less than what there really is)
                    'Shift one down by copying the next item in the array to the current one
                    UnitsDataBase(Count) = UnitsDataBase(Count + 1)
                    For count2 = 1 To ListView1(0).ListItems.Count 'update the key values for the corresponding listitems
                        If (Val(ListView1(0).ListItems(count2).Key) = Count + 1) Then ''The old key will have the count+1 value. Count2 steps through the listview. Count steps through the Database array
                            ListView1(0).ListItems(count2).Key = CStr(Count & "ID")
                            Exit For
                        End If
                    Next
                Next
                
                'Once everything in the unitsdatabase array is shifted down, then the last two items are identical.
                'Redim the array to make it one smaller (thus dropping the last duplicate item)
                'Also, redim only if there is at least one item left
                If (CategoriesUbound >= 0) Then
                    ReDim Preserve UnitsDataBase(CategoriesUbound)
                Else
                    Erase UnitsDataBase
                End If

                'Save the database
                Call SaveUnitsToFile
                
                'Handle the selection aspect (and update the related "Selection" variables)
                'Setup a new selection (the old one was just deleted)
                If (ListView1(0).ListItems.Count - Selection.CategoryListviewRow < 0) Then
                    Selection.CategoryListviewRow = ListView1(0).ListItems.Count
                End If
                
                'If no items were left then don't try to select or recalculate anything
                If (CategoriesUbound < 0) Then
                    Selection.CategoryDatabaseIndex = -1
                    Selection.CategoryUnitsCount = 0
                    ListView1(1).ListItems.Clear
                    Exit Sub
                End If
                
                ListView1(0).ListItems(Selection.CategoryListviewRow).Selected = True
                Selection.CategoryDatabaseIndex = Val(ListView1(0).SelectedItem.Key)
                Selection.CategoryUnitsCount = UnitsDataBase(Selection.CategoryDatabaseIndex).Category.RelatedUnitCount
                
                Call ShowRelatedUnits(True)
                
        '-----------------------------------------------
            ElseIf (ActiveListView = 1) Then 'Delete unit
                Message = "Delete the selected Unit <" & ListView1(1).SelectedItem.Text & "> ?"
                UserDecision = MsgBox(Message, vbOKCancel + vbDefaultButton2, "Confirm")
                If (UserDecision <> 1) Then Exit Sub
                
                'adjust the counter in the database itself
                UnitsDataBase(Selection.CategoryDatabaseIndex).Category.RelatedUnitCount = ListView1(1).ListItems.Count - 1
                
                'Remove the item from the listview (it is important to do it up here, otherwise you will have problems with key assignments)
                ListView1(1).ListItems.Remove (Selection.UnitListviewRow) 'This will be re-indexed automatically
                
                'Shift the database down by one
                'if the deleted item was last on a long list then this for-loop is being bypassed
                For Count = Selection.UnitDatabaseIndex To Selection.CategoryUnitsCount - 2
                    'Shift one down by copying the next item in the array to the current one
                    UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Count) = UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Count + 1)
                    'update the key values for the corresponding listitems
                    For count2 = 1 To ListView1(1).ListItems.Count 'go through the listview items one by one changing their key to a correct one
                        If (Val(ListView1(1).ListItems(count2).Key) = Count + 1) Then 'The old key will have the count+1 value. Count2 steps through the listview. Count steps through the Database array
                            ListView1(1).ListItems(count2).Key = CStr(Count & "ID")
                            Exit For
                        End If
                    Next
                Next
                
                'Once everything in the unitsdatabase array is shifted down, then the last two items are identical.
                'Redim the array to make it one smaller (thus dropping the last duplicate item)
                'Also, redim only if there is at least one item left
                If (UnitsDataBase(Selection.CategoryDatabaseIndex).Category.RelatedUnitCount > 0) Then
                    ReDim Preserve UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(UnitsDataBase(Selection.CategoryDatabaseIndex).Category.RelatedUnitCount - 1)
                Else
                    Erase UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits
                End If
                
                'Save the database
                Call SaveUnitsToFile
                
                'Handle the selection aspect (and update the related "Selection" variables)
                'Setup a new selection (the old one was just deleted)
                If (ListView1(1).ListItems.Count - Selection.UnitListviewRow < 0) Then
                    'if the deleted item was already the last one then make sure it is still the last one now
                    Selection.UnitListviewRow = ListView1(1).ListItems.Count
                End If
                
                Selection.CategoryUnitsCount = UnitsDataBase(Selection.CategoryDatabaseIndex).Category.RelatedUnitCount
                
                'If no items were left then don't try to select or recalculate anything
                If (Selection.CategoryUnitsCount = 0) Then
                    Selection.UnitDatabaseIndex = -1
                    Exit Sub
                End If
                
                ListView1(1).ListItems(Selection.UnitListviewRow).Selected = True
                Selection.UnitDatabaseIndex = Val(ListView1(1).SelectedItem.Key)
                Call FindParent
                Call RecalcValues
            End If
        Case Menu.Edit
            If (ActiveListView = 0) Then
                Call ListView1_Keydown(ActiveListView, 113, 0)
            ElseIf (ActiveListView = 1) Then
                Call EditUnit
            End If
    End Select
End Sub

'-----------------------------------------------------------------------
'GENERAL procedures ----------------------------------------------------
'-----------------------------------------------------------------------

Private Sub PopulateViews()
    Dim Counter As Integer
    
    'Set the headers for the categories list
    ListView1(0).ColumnHeaders.Add , , "Category"
    
    'Set the headers for the units list
    ListView1(1).ColumnHeaders.Add , , "Name", 1350
    ListView1(1).ColumnHeaders.Add , , "Value", 1950, 1
    ListView1(1).ColumnHeaders.Add , , "Symbol", 650
    ListView1(1).ColumnHeaders.Add , , "UnitSystem", 1000, 2
    
    'Exit if there are no categories (the database is empty)
    If (CategoriesUbound < 0) Then Exit Sub
    
    'Populate Categories list
    ListView1(0).Sorted = False
    For Counter = 0 To UBound(UnitsDataBase)
        'index is the listview "Row number" and key is the "array number"
        ListView1(0).ListItems.Add Counter + 1, Counter & "ID", UnitsDataBase(Counter).Category.CategoryName
    Next
    
    ListView1(0).ListItems(1).Selected = True 'Initialize a selection
    
    Call SortList(0)
    Call ShowRelatedUnits(True)
End Sub

Public Sub ShowRelatedUnits(Optional Force As Boolean = False, Optional SelectRow As Integer = 0)
    Dim Counter As Integer
    Dim ConvFactor, Offset As Variant
    Dim ThisExpression As String
      
    'Exit if there are no categories (thus no related units to display)
    If (CategoriesUbound < 0) Then Exit Sub
    Label1(1).Caption = ListView1(0).SelectedItem.Text & " units"
    
    'Exit if the selected category has not changed
    If (Val(ListView1(0).SelectedItem.Key) = Selection.CategoryDatabaseIndex And Force = False) Then Exit Sub
    
    'Remove all rows from the UnitsListview
    ListView1(1).ListItems.Clear
    
    'Update selection variables (to reflect selection of the new category)
    Selection.CategoryDatabaseIndex = Val(ListView1(0).SelectedItem.Key)
    Selection.CategoryListviewRow = ListView1(0).SelectedItem.Index 'the current index (could be different depending on list sorting)
          
    'find how many units are related to the newly selected category
    Selection.CategoryUnitsCount = UnitsDataBase(Selection.CategoryDatabaseIndex).Category.RelatedUnitCount
    If (Selection.CategoryUnitsCount = 0) Then Exit Sub
    
    ParentUnit.Exists = False
    Selection.Exists = False
    ListView1(1).Sorted = False
    For Counter = 0 To Selection.CategoryUnitsCount - 1
        ListView1(1).ListItems.Add Counter + 1, Counter & "ID", UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Counter).LongName
        
        ThisExpression = UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Counter).ConversionFactor
        ConvFactor = EvalExpression(ThisExpression)
        ThisExpression = UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Counter).Offset
        Offset = EvalExpression(ThisExpression)
        
        If (ConvFactor(1) = 0 And Offset(1) = 0) Then
            ListView1(1).ListItems(Counter + 1).ListSubItems.Add 1, , Format(ConvFactor(0) + Offset(0), DecimalRounding)
        ElseIf (ConvFactor(1) <> 0) Then
            ListView1(1).ListItems(Counter + 1).ListSubItems.Add 1, , "Bad Conversion Factor: " & ConvFactor(2)
        ElseIf (Offset(1) <> 0) Then
            ListView1(1).ListItems(Counter + 1).ListSubItems.Add 1, , "Bad Offset Value: " & Offset(2)
        End If
        
        ListView1(1).ListItems(Counter + 1).ListSubItems.Add 2, , UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Counter).ShortName
        ListView1(1).ListItems(Counter + 1).ListSubItems.Add 3, , UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Counter).UnitSystem
        ListView1(1).ListItems(Counter + 1).Tag = Selection.CategoryDatabaseIndex
        ListView1(1).ListItems(Counter + 1).ToolTipText = UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Counter).Description
        
        'Make the parent unit the default selection
        If (Val(ListView1(1).ListItems(Counter + 1).ListSubItems(1).Text) = 1 And SelectRow = 0) Then
            Selection.UnitListviewRow = Counter + 1
            ListView1(1).ListItems(Selection.UnitListviewRow).Selected = True
            Selection.UnitDatabaseIndex = Val(ListView1(1).SelectedItem.Key)
            
        End If
    Next
    
    If (SelectRow > 0) Then 'honor the external row selection request
        Selection.UnitListviewRow = SelectRow
        ListView1(1).ListItems(SelectRow).Selected = True
        Selection.UnitDatabaseIndex = Val(ListView1(1).ListItems(SelectRow).Key)
    End If
    
    Call SortList(1) 'This will also update the parentunit variable
    
    'If no external Row-select request exists and no parent exists, then just select the first item on the list (as long as the sorting is already done)
    If (ParentUnit.Exists = False And SelectRow < 1) Then
        Selection.UnitListviewRow = 1
        ListView1(1).ListItems(Selection.UnitListviewRow).Selected = True
        Selection.UnitDatabaseIndex = Val(ListView1(1).ListItems(1).Key)
    End If
    
    Call RecalcValues
End Sub

Public Sub RecalcValues()
    Dim SelFactor, OtherFactor As Variant
    Dim SelOffset, OtherOffset As Variant
    Dim Count As Integer
    Dim NewValue As Double
    Dim ThisExpression As String
    Dim UserValueForSelUnit As Double
    
    If (Selection.CategoryUnitsCount = 0) Then Exit Sub
    Selection.UnitListviewRow = ListView1(1).SelectedItem.Index 'The current index (could be different depending on list sorting)
    Selection.UnitDatabaseIndex = Val(ListView1(1).SelectedItem.Key)
    
    ThisExpression = UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Selection.UnitDatabaseIndex).ConversionFactor
    SelFactor = EvalExpression(ThisExpression)
    If (SelFactor(1) <> 0 Or SelFactor(0) = 0) Then Exit Sub
    
    ThisExpression = UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Selection.UnitDatabaseIndex).Offset
    SelOffset = EvalExpression(ThisExpression)
    If (SelOffset(1) <> 0) Then Exit Sub
    
    If (UBound(Split(Text1(0).Tag)) >= 0) Then UserValueForSelUnit = Val(Split(Text1(0).Tag)(0)) Else UserValueForSelUnit = Val(Text1(0).Text)
    For Count = 0 To Selection.CategoryUnitsCount - 1
        
        ThisExpression = UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Val(ListView1(1).ListItems(Count + 1).Key)).Offset
        OtherOffset = EvalExpression(ThisExpression)
        ThisExpression = UnitsDataBase(Selection.CategoryDatabaseIndex).RelatedUnits(Val(ListView1(1).ListItems(Count + 1).Key)).ConversionFactor
        OtherFactor = EvalExpression(ThisExpression)
        
        If (OtherFactor(1) = 0 And OtherOffset(1) = 0) Then
            NewValue = (UserValueForSelUnit - SelOffset(0)) * OtherFactor(0) / SelFactor(0) + OtherOffset(0)
            ListView1(1).ListItems(Count + 1).ListSubItems(1).Text = Format(NewValue, DecimalRounding)
        ElseIf (OtherFactor(1) <> 0) Then
            ListView1(1).ListItems(Count + 1).ListSubItems(1).Text = "Bad Conversion factor value: " & OtherFactor(2)
        ElseIf (OtherOffset(1) <> 0) Then
            ListView1(1).ListItems(Count + 1).ListSubItems(1).Text = "Bad Offset value: " & OtherOffset(2)
        End If
        
    Next
    
    ListView1(1).ListItems(Selection.UnitListviewRow).Selected = True
    ListView1(1).SelectedItem.EnsureVisible
End Sub

Private Sub EditUnit()
    Load frmEditUnit
    frmEditUnit.Show
    frmEditUnit.NewEdit(0).SetFocus
End Sub

Public Sub SortList(Index As Integer)
    Dim Count As Integer

    ListView1(Index).Sorted = True
    ListView1(Index).SortKey = 0
    ListView1(Index).Sorted = False
    
    If (ListView1(Index).SelectedItem Is Nothing Or ListView1(Index).ListItems.Count = 0) Then Exit Sub
    Select Case Index
        Case 0
            Selection.CategoryDatabaseIndex = Val(ListView1(0).SelectedItem.Key)
            Selection.CategoryListviewRow = ListView1(0).SelectedItem.Index 'the current index (this could be different depending on list sorting)
            'find the parentUnit again (since things are now suffled)
            Call FindParent
        Case 1
            Selection.UnitListviewRow = ListView1(1).SelectedItem.Index 'The current index (this could be different depending on list sorting)
            Selection.UnitDatabaseIndex = Val(ListView1(1).SelectedItem.Key)
            'find the parentUnit again (since things are now suffled)
            Call FindParent
    End Select
End Sub

Public Sub FindParent()
    Dim Count As Integer
    Dim count2 As Integer
    
    'Asume it doesn't exist
    ParentUnit.Exists = False
    ParentUnit.Count = 0
    ParentUnit.Names = ""
    
    'If lists are empty
    If (ListView1(0).ListItems.Count = 0 Or ListView1(1).ListItems.Count = 0) Then
        Exit Sub
    End If
    
    For Count = 1 To ListView1(1).ListItems.Count
    
        'Listitems color
        'Call ListItemColor(ListView1(1).ListItems(Count), , Color.darkgrey, 2)
        'Call ListItemColor(ListView1(1).ListItems(Count), , Color.black, 0)
        
        If (UnitsDataBase(Val(ListView1(1).ListItems(Count).Tag)).RelatedUnits(Val(ListView1(1).ListItems(Count).Key)).ConversionFactor = 1 And _
            UnitsDataBase(Val(ListView1(1).ListItems(Count).Tag)).RelatedUnits(Val(ListView1(1).ListItems(Count).Key)).Offset = 0) Then
            
            ParentUnit.Count = ParentUnit.Count + 1
            If (ParentUnit.Count > 1) Then ParentUnit.Names = ParentUnit.Names & ", "
            ParentUnit.Names = ParentUnit.Names & UnitsDataBase(Val(ListView1(1).ListItems(Count).Tag)).RelatedUnits(Val(ListView1(1).ListItems(Count).Key)).LongName
            ParentUnit.Exists = True
            ParentUnit.UnitListviewRow = Count
            ParentUnit.UnitDatabaseIndex = Val(ListView1(1).ListItems(Count).Key)
            ParentUnit.CategoryUnitsCount = UnitsDataBase(Val(ListView1(1).ListItems(Count).Tag)).Category.RelatedUnitCount
            ParentUnit.CategoryDatabaseIndex = Val(ListView1(1).ListItems(Count).Tag)
            
            'PrentUnit color
            'Call ListItemColor(ListView1(1).ListItems(Count), , Color.blue, 0)
            
            'Technically the category of the parent is the same as the category of the curent selected unit (but what if there is no currently selected unit ? ... so just find it the long way to be safe)
            For count2 = 1 To ListView1(0).ListItems.Count
                If (Val(ListView1(0).ListItems(count2).Key) = ParentUnit.CategoryDatabaseIndex) Then
                    ParentUnit.CategoryListviewRow = count2
                    Exit For
                End If
            Next
        End If
    Next
End Sub

Public Sub AdjustWindow()
    Text1(1).Left = frmMainWindow.Width - 615 - 270
    Check1.Left = frmMainWindow.Width - 975 - 30 - 615 - 270
    Line1.X2 = frmMainWindow.Width - 300
    Label3(2).Left = frmMainWindow.Width - 495
    Label3(3).Left = frmMainWindow.Width - 720
    Label2(1).Top = frmMainWindow.Height - 780
    Label2(0).Top = frmMainWindow.Height - 780
    Label2(0).Left = frmMainWindow.Width - 3420
    
    Label4.Width = frmMainWindow.Width - Label4.Left - 780
    ListView1(1).Width = frmMainWindow.Width - 2205
    ListView1(1).Height = frmMainWindow.Height - 1620
    ListView1(0).Height = frmMainWindow.Height - 1620
End Sub

'Changes color of Listitem line of a listview control (affects the entire line, or just the headers, or just the subitems, depending on the scope value)
Public Sub ListItemColor(ThisListItem As ListItem, Optional isBold As Boolean = False, Optional textForeColor As Long = -1, Optional Scope As Integer = 0)
    Dim Count As Integer
    
    If (Scope = 0 Or Scope = 1) Then 'All, or Headers only
        ThisListItem.Bold = isBold
        ThisListItem.ForeColor = textForeColor
    End If
    
    If (Scope = 0 Or Scope = 2) Then 'All, or Subitems Only
        For Count = 1 To ThisListItem.ListSubItems.Count
            ThisListItem.ListSubItems(Count).Bold = isBold
            ThisListItem.ListSubItems(Count).ForeColor = textForeColor
        Next
    End If
End Sub

Private Sub FeedbackString(Optional IsWarning As Boolean = True, Optional ThisText As String = "", Optional ThisTag As String = "")
    
    Select Case IsWarning
        Case True
            Label4.ForeColor = RGB(220, 0, 0)
            Label4.Alignment = 0
            Label4.Caption = ThisText
            Label4.Tag = ThisTag
        Case False
            Label4.ForeColor = &H80000011
            Label4.Alignment = 0
            Label4.Caption = ThisText '& String(30 - Len(ListView1(1).HitTest(X, Y).ListSubItems(2).Text), Asc(" "))
            Label4.Tag = ThisTag
    End Select
End Sub
