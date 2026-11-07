VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_UserFormTranspiler 
   Caption         =   "VBA UserForm Transpiler"
   ClientHeight    =   6885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6180
   OleObjectBlob   =   "UI_UserFormTranspiler.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UI_UserFormTranspiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' VBA-UserForm-Transpiler
' https://github.com/ZeeZeX/VBA-UserForm-Transpiler
' Copyright (c) 2026 ZeeZeX
' This software is released under the MIT License.
' https://opensource.org/licenses/MIT


Option Explicit

Private Const OPTION_NAME_TKINTER As String = "Python(Tkinter)"
Private Const OPTION_NAME_PS_WINFORMS As String = "PowerShell(WinForms)"
Private Const OPTION_NAME_HTML As String = "HTML/CSS"


Private Sub btnConvert_Click()
    
    Dim wb As Workbook
    Dim uf As Object
    
    If cbUserForms.value = "" Then
        MsgBox "No UserForm selected."
        Exit Sub
    End If
    
    If Not Contains(GetWorkbookNames(), cbWorkbooks.value) Then
        MsgBox "The selected workbook was not found."
        Exit Sub
    End If
    
    Set wb = Workbooks(cbWorkbooks.value)
    
    If Not Contains(GetUserFormNames(wb), cbUserForms.value) Then
        MsgBox "The selected UserForm was not found."
        Exit Sub
    End If
    
    Set uf = GetUserFormFromWorkbook(wb, cbUserForms.value)
    uf.Hide
    Select Case cbTarget.value
        Case OPTION_NAME_TKINTER
            Call ConvertForm2Tkinter(uf, _
                useCls:=chkUseClass.value, _
                noMainLoop:=chkNoMainLoop.value, _
                uniqueStyleName:=chkUniqueStyleName.value, _
                imageMode:=cbImageMode.value)
                
        Case OPTION_NAME_PS_WINFORMS
            Call ConvertForm2PS(uf, _
                saveAsBat:=chkSaveAsBat.value, _
                useCls:=chkUseClass.value, _
                noMainLoop:=chkNoMainLoop.value, _
                imageMode:=cbImageMode.value)
        Case OPTION_NAME_HTML
            Call ConvertForm2HTML(uf, _
                usePrefix:=chkUseClass.value, _
                imageMode:=cbImageMode.value, _
                langAttribute:=cbLangAttribute)
        Case Else
            MsgBox "Failed: Invalid Target"
    End Select
End Sub

Private Sub cbTarget_Change()
    Call SetOptionControlsPosition
    Call ListImageMode
End Sub


Private Sub UserForm_Initialize()
    cbTarget.Clear
    cbTarget.List = Array(OPTION_NAME_TKINTER, OPTION_NAME_PS_WINFORMS, OPTION_NAME_HTML)
    cbTarget.value = cbTarget.List(0)
    Call ResetComboBox
    
    With cbLangAttribute
        .List = Array("", "en", "en-US", "zh", "zh-CN", "zh-TW", "ja")
    End With
End Sub

Private Sub SetOptionControlsPosition()
    Dim baseTop As Long
    Dim baseLeft As Long
    Dim offset As Long
    Dim allCtrls() As Variant
    Dim availableCtrls() As Variant
    Dim ctrl As Variant
    baseTop = labelOptions.Top + labelOptions.height + 5
    baseLeft = labelOptions.Left
    
    
    Select Case cbTarget.value
        Case OPTION_NAME_TKINTER
            availableCtrls = Array(chkUseClass, chkNoMainLoop, chkUniqueStyleName, frameImageMode)
        Case OPTION_NAME_PS_WINFORMS
            availableCtrls = Array(chkUseClass, chkSaveAsBat, chkNoMainLoop, frameImageMode)
        Case OPTION_NAME_HTML
            availableCtrls = Array(chkUseClass, frameLangAttribute, frameImageMode)
        Case Else
            availableCtrls = Array()
    End Select
    
    allCtrls = Array(chkUseClass, chkNoMainLoop, chkUniqueStyleName, chkSaveAsBat, frameLangAttribute, frameImageMode)
    For Each ctrl In allCtrls
        ctrl.Visible = False
    Next ctrl
    
    offset = 0
    For Each ctrl In availableCtrls
        ctrl.Visible = True
        ctrl.Top = baseTop + offset
        ctrl.Left = baseLeft
        offset = offset + ctrl.height + 5
    Next ctrl
    
End Sub

Private Sub ListUserForms(ByVal wb As Workbook)
    Dim arrUserForms() As Variant
    cbUserForms.Clear
    arrUserForms = GetUserFormNames(wb)
    If ArrayCount(arrUserForms) = 0 Then
        cbUserForms.value = ""
    Else
        cbUserForms.List = arrUserForms
        cbUserForms.value = arrUserForms(0)
    End If
    
End Sub

Private Sub ListWorkbooks()
    Dim arrWorkbookNames() As Variant
    cbWorkbooks.Clear
    
    arrWorkbookNames = GetWorkbookNames()
    cbWorkbooks.List = arrWorkbookNames
    If ArrayCount(arrWorkbookNames) >= 2 Then
        If arrWorkbookNames(0) = ThisWorkbook.Name Then
            cbWorkbooks.value = arrWorkbookNames(1)
        Else
            cbWorkbooks.value = arrWorkbookNames(0)
        End If
    ElseIf ArrayCount(arrWorkbookNames) >= 1 Then
        cbWorkbooks.value = arrWorkbookNames(0)
    Else
        cbWorkbooks.value = ""
    End If
    
End Sub

Private Sub ListImageMode()
    Dim selectableOptions() As Variant
    Dim previousValue As String
    Select Case cbTarget.value
        Case OPTION_NAME_TKINTER
            selectableOptions = Array("file", "disabled", "reference-only", "base64", "base64-dict", "base64-json", "base64-json-reference")
        Case OPTION_NAME_PS_WINFORMS
            selectableOptions = Array("file", "disabled", "reference-only", "base64", "base64-dict", "base64-json", "base64-json-reference")
        Case OPTION_NAME_HTML
            selectableOptions = Array("file", "disabled", "reference-only", "base64")
        Case Else
            selectableOptions = Array("")
    End Select
    
    With cbImageMode
        previousValue = .value
        .Clear
        .List = selectableOptions
        If Contains(selectableOptions, previousValue) Then
            .value = previousValue
        Else
            .value = .List(0)
        End If
    End With
End Sub

Private Function GetUserFormNames(ByVal wb As Workbook) As Variant()
    Dim vbComp As Object
    Dim collUserForms As Collection
    Dim arrUserForms() As Variant
    Set collUserForms = New Collection
    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Type = 3 Then  ' vbext_ct_MSForm
            collUserForms.Add vbComp.Name
        End If
    Next vbComp
    arrUserForms = CollectionToArray(collUserForms)
    If collUserForms.count > 0 Then
        Call QuickSortArray(arrUserForms)
    End If
    GetUserFormNames = arrUserForms
End Function

Private Function GetWorkbookNames() As Variant()
    Dim wb As Workbook
    Dim arrWorkbookNames() As Variant
    Dim collWorkbookNames As Collection
    Set collWorkbookNames = New Collection
    For Each wb In Workbooks
        collWorkbookNames.Add wb.Name
    Next
    arrWorkbookNames = CollectionToArray(collWorkbookNames)
    GetWorkbookNames = arrWorkbookNames
End Function


Private Function Contains(ByRef arr As Variant, ByVal targetValue As Variant) As Boolean
    Dim item As Variant
    For Each item In arr
        If item = targetValue Then
            Contains = True
            Exit Function
        End If
    Next
    Contains = False
End Function

Private Function ArrayCount(ByRef arr As Variant) As Long
    ArrayCount = UBound(arr) - LBound(arr) + 1
End Function

Private Sub ResetComboBox()
    Call ListWorkbooks
    Call ListUserForms(Workbooks(cbWorkbooks.value))
    Call ListImageMode
End Sub


Private Sub btnReload_Click()
    Call ResetComboBox
End Sub

Private Sub cbWorkbooks_Change()
    If cbWorkbooks.value <> "" Then Call ListUserForms(Workbooks(cbWorkbooks.value))
End Sub

Private Function GetUserFormFromWorkbook( _
    ByVal wb As Workbook, _
    ByVal formName As String) As Object

    Dim isSaved As Boolean
    Dim comp As Object
    Dim tmpComp As Object
    Dim frm As Object
    Dim result As Object
    Dim moduleName As String
    Dim codeText As String
    Dim i As Long

    Dim errNum As Long
    Dim errDesc As String
    
    isSaved = wb.Saved
    
    If wb Is ThisWorkbook Then
        Set GetUserFormFromWorkbook = VBA.UserForms.Add(formName)
        Exit Function
    End If
    
    On Error GoTo ErrorHandler

    ' Verify that the specified component exists
    Set comp = wb.VBProject.VBComponents(formName)

    If comp.Type <> 3 Then
        Err.Raise vbObjectError + 1000, , _
                  "The specified component is not a UserForm."
    End If

    ' Determine a unique temporary module name.
    ' Naming convention: "zzTmpUF" followed by a zero-padded
    ' three-digit sequential number (e.g., zzTmpUF001).
    ' Existing module names are skipped to avoid naming conflicts.
    Do
        i = i + 1
        moduleName = "zzTmpUF" & format$(i, "000")
        Set comp = Nothing

        On Error Resume Next
        Set comp = wb.VBProject.VBComponents(moduleName)
        On Error GoTo ErrorHandler

        If comp Is Nothing Then Exit Do
    Loop

    ' Add a temporary standard module to the target workbook
    Set tmpComp = wb.VBProject.VBComponents.Add(1)
    tmpComp.Name = moduleName

    codeText = _
        "Public Function CreateForm(ByVal nm As String) As Object" & vbCrLf & _
        "    Dim f As Object" & vbCrLf & _
        "    Set f = VBA.UserForms.Add(nm)" & vbCrLf & _
        "    f.Show vbModeless" & vbCrLf & _
        "    Set CreateForm = f" & vbCrLf & _
        "End Function"

    tmpComp.CodeModule.AddFromString codeText

    ' Retrieve the UserForm object as the function return value
    Set result = Application.Run( _
        "'" & Replace(wb.Name, "'", "''") & "'!" & _
        moduleName & ".CreateForm", formName)

    Set GetUserFormFromWorkbook = result

CleanUp:
    On Error Resume Next
    If Not tmpComp Is Nothing Then
        wb.VBProject.VBComponents.Remove tmpComp
    End If
    On Error GoTo 0

    wb.Saved = isSaved
    
    If errNum <> 0 Then
        Err.Raise errNum, , errDesc
    End If
    Exit Function

ErrorHandler:
    errNum = Err.Number
    errDesc = Err.Description
    Resume CleanUp

End Function


Private Sub QuickSortArray(ByRef vArray As Variant)
    If UBound(vArray) - LBound(vArray) + 1 = 0 Then
        Exit Sub
    End If
    Call QuickSortHelper(vArray, LBound(vArray), UBound(vArray))
End Sub

Private Sub QuickSortHelper(ByRef vArray As Variant, ByVal inLow As Long, ByVal inHi As Long)
    Dim pivot As Variant
    Dim tmpSwap As Variant
    Dim tmpLow As Long
    Dim tmpHi As Long

    tmpLow = inLow
    tmpHi = inHi
    
    pivot = vArray((inLow + inHi) \ 2)

    Do While (tmpLow <= tmpHi)
        Do While (vArray(tmpLow) < pivot) And (tmpLow < inHi)
            tmpLow = tmpLow + 1
        Loop

        Do While (pivot < vArray(tmpHi)) And (tmpHi > inLow)
            tmpHi = tmpHi - 1
        Loop

        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Loop

    If (inLow < tmpHi) Then QuickSortHelper vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSortHelper vArray, tmpLow, inHi
End Sub

Private Function CollectionToArray(ByVal coll As Collection, Optional ByVal isStartIdx1 As Boolean = False) As Variant()
    ' Convert a Collection to an array
    ' If isStartIdx1 is True, create an array starting from index 1 (to match Collection numbering)
    Dim arr() As Variant
    Dim item As Variant
    Dim idx As Long
    If coll.count > 0 Then
        If isStartIdx1 Then
            ReDim arr(1 To coll.count)
        Else
            ReDim arr(0 To coll.count - 1)
        End If
        idx = LBound(arr)
        For Each item In coll
            ' Use "Set" when assigning objects.
            If IsObject(item) Then
                Set arr(idx) = item
            Else
                arr(idx) = item
            End If
            idx = idx + 1
        Next
    Else
        arr = VBA.Array()
    End If
    CollectionToArray = arr
End Function




