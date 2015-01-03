Attribute VB_Name = "Info"
Public Type UnitCategory
    CategoryID As Integer
    CategoryName As Variant
    RelatedUnitCount As Integer
End Type

Public Type Units
    Serial As Integer
    LinkedToCat As Integer
    LongName As Variant
    ShortName As Variant
    ConversionFactor As Variant
    Offset As Variant
    UnitSystem As String
    Description As String
End Type

Public Type UnitFolder
    Category As UnitCategory
    RelatedUnits() As Units
End Type

Public Type UnitInfo
    Names As String
    Exists As Boolean 'This is getting obsolete ... just use the "count"
    Count As Integer
    CategoryUnitsCount As Integer
    CategoryDatabaseIndex As Integer
    CategoryListviewRow As Integer
    UnitListviewRow As Integer
    UnitDatabaseIndex As Integer
End Type

Public Type ExpSyntax
    LastKeyStroke As Integer
    LastKeyCode As Integer
    LastUsefulNum As Double
    OpenParenCount As Integer
    CloseParenCount As Integer
    PoorParenPair As Integer
    InsertionPoint As Integer
    AsciiBeforeInsPoint As Integer
    AsciiAfterInsPoint As Integer
End Type

Public Type NumberInfo
    Number As Double
    DigitsBeforeDecim As Integer
    DigitsAfterDecim As Integer
    DecimalCount As Integer
End Type

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long ''This is to pin the window on top (or not)
Declare Function GetForegroundWindow Lib "user32" () As Long 'This is to see which window has keyboard focus

Public CategoriesUbound As Integer
Public UnitsDataBase() As UnitFolder
Public Selection As UnitInfo
Public ParentUnit As UnitInfo
Public EditMode As Boolean
Public DecimalRounding As String
Public TrackExpSyntax As ExpSyntax
