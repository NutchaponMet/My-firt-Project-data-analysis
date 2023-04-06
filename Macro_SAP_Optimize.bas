Attribute VB_Name = "Module5"
Sub Macro_Complet_Optimize()
Attribute Macro_Complete.VB_ProcData.VB_Invoke_Func = "M\n14"
'
' VBA Macro for organization SAP DATA
'
'
' Keyboard Shortcut: Ctrl+Shift+M
'
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("F2:F205") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A1:AK205")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Range( _
        "A:A,C:C,D:D,H:H,K:K,L:L,N:N,R:R,T:T,V:V,W:W,X:X,Y:Y,Z:Z,AB:AB,AC:AC,AD:AD,AE:AE,AF:AF,AK:AK" _
        ).Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("L:L").Select
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight

' Create Variable

    Dim k As Integer
    Dim i As Integer
    
' Create array
    Arrayvalues = Array("HM1", "HMB", "HML", "HMS") 'Change Hub Name
    k = 0

' Keyboard Shortcut: Ctrl+Shift+O
' Write Relative Reference style
    
    Range("A1").Select
    ActiveCell.Rows("1:3").EntireRow.Select
    ActiveCell.Activate
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(1, 0).Range("A1").Select
    
' Loop 1
    For i = 1 To 4
        Cells.Find(What:=Arrayvalues(k), After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
            , SearchFormat:=False).Activate
        ActiveCell.Rows("1:3").EntireRow.Select
        ActiveCell.Activate
        Selection.Insert Shift:=xlDown
        ActiveCell.Offset(1, 0).Range("A1").Select
        k = k + 1
    Next i
    
' loop2
    Dim v As Integer
    Arrayvalues2 = Array("HK1", "HM1", "HMB", "HML", "HMS") 'Change Hub Name
    v = 0
    Range("A1").Select
    For i = 1 To 5
        Range("A1").Select
        Cells.Find(What:=Arrayvalues2(v), After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
                    , SearchFormat:=False).Activate
        ActiveCell.Offset(0, -1).Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo, _
                       Key2:=Range("E1"), Order1:=xlAscending, Header:=xlNo, _
                       Key3:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        v = v + 1
    Next i
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Cells.Select
    Cells.EntireColumn.AutoFit
    
End Sub

