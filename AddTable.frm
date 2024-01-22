VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddTable 
   Caption         =   "Добавить таблицу"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7575
   OleObjectBlob   =   "AddTable v2.0.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Vesrion 2.0 16.10.2023

Dim objExclApp As Excel.Application
Dim objExclDoc As Excel.Workbook



'Выбор файла эксель
Private Sub ButtonLoadExcel_Click()
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        If .Show = -1 Then
            TextBoxAdrExcel.Text = .SelectedItems(1)
        Else
            .Show
        End If
    End With
    
    If Not TextBoxAdrExcel.Text = "" Then
        ComboBoxExcelShet.Clear
    
        Set objExclApp = New Excel.Application
        Set objExclDoc = objExclApp.Workbooks.Open(TextBoxAdrExcel.Text)
        For i = 1 To objExclDoc.Sheets.Count
            ComboBoxExcelShet.AddItem objExclDoc.Sheets(i).Name
        Next i
        
        TextBoxAdrExcel.Visible = True
        ComboBoxExcelShet.Visible = True
        LabelAdrExcel.Visible = True
        
        objExclApp.Visible = False
        objExclDoc.Close False
        objExclApp.Quit
        
        Set objExclApp = Nothing
        Set objExclDoc = Nothing
    End If
End Sub


'Добавление столбцов из эксель
Private Sub ComboBoxExcelShet_Change()
    Set objExclApp = New Excel.Application
    Set objExclDoc = objExclApp.Workbooks.Open(TextBoxAdrExcel.Text)
    
    ListBoxExcel.Clear
    ListBoxWord.Clear
    ComboBoxColFilter.Clear
    ListBoxNameFilter.Clear
    ListBoxFiltersTable.Clear
    
    ButtonAddColl.Enabled = False
    ButtonDelColl.Enabled = False
    ButtonAddFilter.Visible = False
    ButtonDelFilter.Visible = False
    ButtonAddTable.Enabled = False

   
    ListBoxExcel.AddItem
    ListBoxExcel.column(0, 0) = 0
    ListBoxExcel.column(1, 0) = "№ п/п"
    For intColsExcel = 1 To objExclDoc.Sheets(ComboBoxExcelShet.Value).Cells(1, objExclDoc.Sheets(ComboBoxExcelShet.Value).Columns.Count).End(xlToLeft).column
        If "" = objExclDoc.Sheets(ComboBoxExcelShet.Value).Cells(1, intColsExcel).Value Then
            ListBoxExcel.AddItem
            ListBoxExcel.column(1, intColsExcel) = "ПУСТО"
            ListBoxExcel.column(0, intColsExcel) = Split(objExclDoc.Sheets(ComboBoxExcelShet.Value).Cells(, intColsExcel).Address, "$")(1)
            ComboBoxColFilter.AddItem
            ComboBoxColFilter.column(1, intColsExcel - 1) = "ПУСТО"
            ComboBoxColFilter.column(0, intColsExcel - 1) = Split(objExclDoc.Sheets(ComboBoxExcelShet.Value).Cells(, intColsExcel).Address, "$")(1)
        Else
            ListBoxExcel.AddItem
            ListBoxExcel.column(1, intColsExcel) = objExclDoc.Sheets(ComboBoxExcelShet.Value).Cells(1, intColsExcel).Value
            ListBoxExcel.column(0, intColsExcel) = Split(objExclDoc.Sheets(ComboBoxExcelShet.Value).Cells(, intColsExcel).Address, "$")(1)
            ComboBoxColFilter.AddItem
            ComboBoxColFilter.column(1, intColsExcel - 1) = objExclDoc.Sheets(ComboBoxExcelShet.Value).Cells(1, intColsExcel).Value
            ComboBoxColFilter.column(0, intColsExcel - 1) = Split(objExclDoc.Sheets(ComboBoxExcelShet.Value).Cells(, intColsExcel).Address, "$")(1)
        End If
    Next intColsExcel
    
    TextBoxNameColWord.Visible = True
    
    ListBoxExcel.Visible = True
    ListBoxWord.Visible = True
    ListBoxFiltersTable.Visible = True
    ListBoxNameFilter.Visible = True
    
    ButtonAddColl.Visible = True
    ButtonDelColl.Visible = True
    ButtonAddFilter.Visible = True
    ButtonDelFilter.Visible = True
    ButtonAddTable.Visible = True
    
    LabelColunmsExcel.Visible = True
    LabelColunmsWord.Visible = True
    LabelNameColunmWord.Visible = True
    LabelFilter.Visible = True
    
    ComboBoxColFilter.Visible = True
  
    
    objExclApp.Visible = False
    objExclDoc.Close False
    objExclApp.Quit
    
    Set objExclApp = Nothing
    Set objExclDoc = Nothing
End Sub

'Выбор столбцов в эксель
Private Sub ListBoxExcel_Click()
    ButtonAddColl.Enabled = True
    ButtonDelColl.Enabled = False
    TextBoxNameColWord.Value = ListBoxExcel.List(ListBoxExcel.ListIndex, 1)
    
End Sub


'Выбор столбцов для ворд
Private Sub ListBoxWord_Click()
    ButtonAddColl.Enabled = False
    ButtonDelColl.Enabled = True
End Sub

'Добавление столбцов для ворд
Private Sub ButtonAddColl_Click()
    If ListBoxExcel.ListIndex = -1 Then
        MsgBox "Не выбрано"
        ButtonAddTable.Enabled = False
    Else
        ListBoxWord.AddItem
        ListBoxWord.column(0, ListBoxWord.ListCount - 1) = ListBoxExcel.List(ListBoxExcel.ListIndex, 0)
        ListBoxWord.column(1, ListBoxWord.ListCount - 1) = ListBoxExcel.List(ListBoxExcel.ListIndex, 1)
        ListBoxWord.column(2, ListBoxWord.ListCount - 1) = TextBoxNameColWord.Value
        
        ButtonAddTable.Enabled = True
    End If
End Sub

'Удаление столбцов столбцов для ворд
Private Sub ButtonDelColl_Click()
    If ListBoxWord.ListIndex = -1 Then
        MsgBox "Не выбрано"
        ButtonDelColl.Enabled = False
    Else
        ListBoxWord.RemoveItem (ListBoxWord.ListIndex)
    End If
    
    If ListBoxWord.ListCount = 0 Then
        ButtonDelColl.Enabled = False
        ButtonAddTable.Enabled = False
    End If
End Sub

'Добавить № п/п
Private Sub CheckBoxNumber_Click()
    If CheckBoxNumber.Value = True Then
        ListBoxWord.AddItem index = 0
        ListBoxWord.column(0, 0) = 0
        ListBoxWord.column(1, 0) = "№ п/п"
        ListBoxWord.column(2, 0) = "№ п/п"
    Else
        ListBoxWord.RemoveItem (0)
    End If
End Sub


'Выбор столбца для фильтрации
Private Sub ComboBoxColFilter_Change()
    If Not ComboBoxColFilter.ListIndex = -1 Then
        ListBoxNameFilter.Clear
        Set objExclApp = New Excel.Application
        Set objExclDoc = objExclApp.Workbooks.Open(TextBoxAdrExcel.Text)
         
        objExclDoc.Sheets(ComboBoxExcelShet.Value).Cells.Rows.Hidden = False
        objExclDoc.Sheets(ComboBoxExcelShet.Value).Range(objExclDoc.Sheets(ComboBoxExcelShet.Value).Cells(1, 1), _
                                 objExclDoc.Sheets(ComboBoxExcelShet.Value).Cells(objExclDoc.Sheets(ComboBoxExcelShet.Value).Cells(objExclDoc.Sheets(ComboBoxExcelShet.Value).Rows.Count, ComboBoxColFilter.ListIndex + 1).End(xlUp).Row, ComboBoxColFilter.ListIndex + 1)).RemoveDuplicates Columns:=ComboBoxColFilter.ListIndex + 1, Header:=xlYes
         
        For intRowsExcel = 2 To objExclDoc.Sheets(ComboBoxExcelShet.Value).Cells(objExclDoc.Sheets(ComboBoxExcelShet.Value).Rows.Count, ComboBoxColFilter.ListIndex + 1).End(xlUp).Row
            ListBoxNameFilter.AddItem objExclDoc.Sheets(ComboBoxExcelShet.Value).Cells(intRowsExcel, ComboBoxColFilter.ListIndex + 1).Value
        Next
        
        objExclApp.Visible = False
        objExclDoc.Close False
        objExclApp.Quit
         
        Set objExclApp = Nothing
        Set objExclDoc = Nothing
    End If
End Sub

'Выбор фильтра столбца для ворд
Private Sub ListBoxNameFilter_Change()
    ButtonAddFilter.Enabled = True
    ButtonDelFilter.Enabled = False
End Sub

'Выбор фильтра

Private Sub ListBoxFiltersTable_Click()
    ButtonAddFilter.Enabled = False
    ButtonDelFilter.Enabled = True
End Sub

Private Sub ListBoxNameFilter_Click()
    ButtonAddFilter.Enabled = True
    ButtonDelFilter.Enabled = False
End Sub

'Добавление фильтра
Private Sub ButtonAddFilter_Click()
    If ListBoxNameFilter.ListIndex = -1 Then
        MsgBox "Не выбрано"
    Else
        Dim ArrayNames As String

        ArrayNames = ""
        For Count = 1 To ListBoxNameFilter.ListCount
            If ListBoxNameFilter.Selected(Count - 1) Then
                If ArrayNames = "" Then
                    ArrayNames = ListBoxNameFilter.List(Count - 1)
                Else
                    ArrayNames = ArrayNames & " | " & ListBoxNameFilter.List(Count - 1)
                End If
            End If
        Next
        
        ListBoxFiltersTable.AddItem
        ListBoxFiltersTable.column(0, ListBoxFiltersTable.ListCount - 1) = ComboBoxColFilter.List(ComboBoxColFilter.ListIndex, 0)
        ListBoxFiltersTable.column(1, ListBoxFiltersTable.ListCount - 1) = ComboBoxColFilter.List(ComboBoxColFilter.ListIndex, 1)
        ListBoxFiltersTable.column(2, ListBoxFiltersTable.ListCount - 1) = ArrayNames
    End If
End Sub

'Удаление фильтра
Private Sub ButtonDelFilter_Click()
    If ListBoxFiltersTable.ListIndex = -1 Then
        MsgBox "Не выбрано"
    Else
        ListBoxFiltersTable.RemoveItem (ListBoxFiltersTable.ListIndex)
        If ListBoxFiltersTable.ListCount = 0 Then
            ButtonDelFilter.Enabled = False
        End If
    End If
End Sub

'Создание таблицы
Private Sub ButtonAddTable_Click()
    If ListBoxWord.ListCount = 0 Then
        MsgBox "Не выбраны столбцы"
    Else
        Set objExclApp = New Excel.Application
        Set objExclDoc = objExclApp.Workbooks.Open(TextBoxAdrExcel.Text, False, True)
        
        'Фильтрация
        objExclDoc.Sheets(ComboBoxExcelShet.Value).Cells.Rows.Hidden = False
        For Count = 1 To ListBoxFiltersTable.ListCount
            Dim VbaArrayNames() As String
            VbaArrayNames = Split(ListBoxFiltersTable.List(Count - 1, 2), " | ")
            objExclDoc.Sheets(ComboBoxExcelShet.Value).Range("A1").AutoFilter Field:=objExclDoc.Sheets(ComboBoxExcelShet.Value).Range(ListBoxFiltersTable.List(Count - 1, 0) & 1).column, Criteria1:=VbaArrayNames, Operator:=xlFilterValues
        Next
        
        'Создание таблицы
        Dim docActive As Document
        Set docActive = ActiveDocument
        objExclDoc.Sheets.Add.Name = "toWord"
        Dim maxRows As Long
        For Count = 1 To ListBoxWord.ListCount
            If ListBoxWord.List(Count - 1, 0) = 0 Then
                objExclDoc.Sheets("toWord").Cells(1, Count).Value = "№ п/п"
            Else
                objExclDoc.Sheets(ComboBoxExcelShet.Value).Columns(objExclDoc.Sheets(ComboBoxExcelShet.Value).Range(ListBoxWord.List(Count - 1, 0) & 1).column).SpecialCells(xlVisible).Copy
                objExclDoc.Sheets("toWord").Columns(Count).PasteSpecial Paste:=xlPasteFormats
                objExclDoc.Sheets("toWord").Columns(Count).PasteSpecial Paste:=xlPasteValues
                objExclDoc.Sheets("toWord").Cells(1, Count).Value = ListBoxWord.List(Count - 1, 2)
                
                If maxRows < objExclDoc.Sheets("toWord").Cells(objExclDoc.Sheets("toWord").Rows.Count, Count).End(xlUp).Row Then
                    maxRows = objExclDoc.Sheets("toWord").Cells(objExclDoc.Sheets("toWord").Rows.Count, Count).End(xlUp).Row
                End If
            End If
        Next
        
        For Count = 1 To ListBoxWord.ListCount
            If ListBoxWord.List(Count - 1, 0) = 0 Then
                objExclDoc.Sheets("toWord").Cells(2, Count).Value = 1
                If maxRows > 2 Then
                    AdrCol = Split(objExclDoc.Sheets(ComboBoxExcelShet.Value).Cells(, Count).Address, "$")(1)
                    objExclDoc.Sheets("toWord").Range(AdrCol & "2").AutoFill Destination:=objExclDoc.Sheets("toWord").Range(AdrCol & "2:" & AdrCol & maxRows), Type:=xlFillSeries
                End If
                Exit For
            End If
        Next
        objExclDoc.Sheets("toWord").Range("A1:" & Split(objExclDoc.Sheets("toWord").Cells(, ListBoxWord.ListCount - 1).Address, "$")(1) & maxRows).Borders.LineStyle = xlContinuous
        
        
        objExclApp.CutCopyMode = False
        
        'Копирование в ворд
        objExclDoc.Sheets("toWord").Range(objExclDoc.Sheets("toWord").Cells(1, 1), _
        objExclDoc.Sheets("toWord").Cells(objExclDoc.Sheets("toWord").Cells(objExclDoc.Sheets("toWord").Rows.Count, 1).End(xlUp).Row, ListBoxWord.ListCount)).CurrentRegion.Copy
        Selection.PasteExcelTable False, True, True
        objExclApp.CutCopyMode = False
        
        'Форматирование в ворд
        Dim t As Table
        For Each t In ActiveDocument.Tables
            If t.Columns.Count = ListBoxWord.ListCount And t.Rows.Count = objExclApp.WorksheetFunction.Subtotal(3, _
                            objExclDoc.Sheets("toWord").Range(objExclDoc.Sheets("toWord").Cells(2, 1), _
                            objExclDoc.Sheets("toWord").Cells(objExclDoc.Sheets("toWord").Cells(objExclDoc.Sheets("toWord").Rows.Count, 1).End(xlUp).Row, 1))) + 1 Then
                With t
                    .Style = "Сетка таблицы"
                    .Range.Font.Name = "Arial"
                    For Count = 1 To ListBoxWord.ListCount
                        .Cell(1, Count).Range.Bold = True
                        .Cell(1, Count).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        .Rows(1).HeadingFormat = True
                    Next
                    .PreferredWidthType = wdPreferredWidthPercent
                    .PreferredWidth = 100
                    .Columns.AutoFit
                End With
            End If
        Next
        
        objExclApp.Visible = True
        objExclDoc.Close False
        objExclApp.Quit
    
        Set objExclApp = Nothing
        Set objExclDoc = Nothing
    End If
    Unload AddTable
End Sub
