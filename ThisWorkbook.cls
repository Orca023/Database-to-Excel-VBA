VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' Excel 文檔打開時的操作
Private Sub Workbook_Open()

    On Error Resume Next

    ''遍歷所有模塊，並移除不符合條件的已存在模塊
    'Dim moduleName
    'moduleName = Array("clsBrowser", "clsCore", "clsJsConverter", "DatabaseDispatchModule", "DatabaseControlPanel", "DatabaseModule")
    'Dim i As Integer, j As Integer
    'Dim flagBoolean As Boolean
    'flagBoolean = True
    'For i = 1 To ThisWorkbook.VBProject.VBComponents.Count
    '    flagBoolean = True
    '    For j = LBound(moduleName) To UBound(moduleName)
    '        If CStr(ThisWorkbook.VBProject.VBComponents(i).Name) = moduleName(j) Then
    '            flagBoolean = False
    '            Exit For
    '        End If
    '    Next j
    '    If (CStr(ThisWorkbook.VBProject.VBComponents(i).Name) = "ThisWorkbook") Then
    '        flagBoolean = False
    '    End If
    '    For j = 1 To ThisWorkbook.Worksheets.Count
    '        If CStr(ThisWorkbook.VBProject.VBComponents(i).Name) = CStr(ThisWorkbook.Worksheets(j).Name) Then
    '            flagBoolean = False
    '            Exit For
    '        End If
    '    Next j
    '    If flagBoolean Then
    '        'Debug.Print ThisWorkbook.VBProject.VBComponents(i).Name
    '        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(i): Rem 移除模塊
    '    End If
    'Next i

    'ThisWorkbook.Save

    Call DatabaseDispatchModule.initial  '調用鏈接操控數據庫調度模組的自定義子過程變量初始化賦初值
    Call DatabaseDispatchModule.MenuSetup  '調用鏈接操控數據庫調度模組的自定義子過程創建自定義菜單欄

End Sub


' Excel 文檔關閉前的操作
Private Sub Workbook_BeforeClose(Cancel As Boolean)

    On Error Resume Next

    'Call DatabaseDispatchModule.callRemoveModule  '調用鏈接操控數據庫調度模組的自定義子過程移除清理程序自動導入的模塊
    ''Call DatabaseDispatchModule.callKillProcess  '調用鏈接操控數據庫調度模組的自定義子過程關閉清理程序自動創建的進程

    'ThisWorkbook.Save

End Sub


''導入指定的模塊文檔，並自定義重命名導入的模塊
'Public Sub importModule(moduleFile As String, moduleName As String)
'
'    On Error Resume Next

'    ''自定義的待導入的統計算法模塊
'    'Dim moduleName As String, modulePath As String, moduleFileName As String, moduleFile As String
'    'moduleName = "KLineChart"  '鏈接操控 MongoDB 數據庫模塊的自定義命名
'    'modulePath = "C:\Criss\vba\QualityControl\KLineChart"  '鏈接操控 MongoDB 數據庫模塊的保存路徑
'    'moduleFileName = "KLineChart.bas"  '鏈接操控 MongoDB 數據庫模塊的文檔名
'    'moduleFile = modulePath & "\" & moduleFileName  '鏈接操控 MongoDB 數據庫模塊文檔的路徑全名

'    '判斷自定義的統計算法模塊是否保存在指定的路徑
'    Dim FSO As Object
'    Set FSO = CreateObject("Scripting.FileSystemObject")
'    'If fso.Folderexists(modulePath) And fso.Fileexists(moduleFile) Then
'    If FSO.FileExists(moduleFile) Then

'        '使用 Is Nothing 方法判斷指定名字的模塊是否已經存在，注意需要寫上語句 On Error Resume Next，不然當沒有找到指定的模塊時會報錯
'        If ThisWorkbook.VBProject.VBComponents(moduleName) Is Nothing Then

'            '從硬盤文檔導入模塊
'            With ThisWorkbook.VBProject.VBComponents.Import(moduleFile)
'                .name = moduleName  '對導入的模塊自定義重命名
'            End With

'        Else

'            '首先移除指定模塊 moduleName
'            ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(moduleName)  '移除指定模塊 moduleName
'            '導出指定模塊 moduleName 至指定文檔 moduleFile
'            'ThisWorkbook.VBProject.VBComponents(moduleName).Export moduleFile
'            '模塊重命名
'            'ThisWorkbook.VBProject.VBComponents("A").Name = "B"

'            '然後再次使用 Is Nothing 方法判斷指定名字的模塊是否已經存在，注意需要寫上語句 On Error Resume Next，不然當沒有找到指定的模塊時會報錯
'            If ThisWorkbook.VBProject.VBComponents(moduleName) Is Nothing Then

'                '從硬盤文檔導入模塊
'                With ThisWorkbook.VBProject.VBComponents.Import(moduleFile)
'                    .name = moduleName  '對導入的模塊自定義重命名
'                End With

'            Else

'                Debug.Print "imported Module ( " & moduleName & " ) error, Module ( " & moduleName & " ) already exists, and unable to remove update."
'                'Exit Sub

'            End If

'        End If

'        'Debug.Print "import Module ( " & moduleName & " ) source file: " & moduleFile

'    Else

'        Debug.Print "import Module ( " & moduleName & " ) error, Source file ( " & moduleFile & " ) is Nothing."

'    End If

'    Set FSO = Nothing

'End Sub


''移除指定的已存在模塊
'Public Sub removeModule(moduleName As String)

'    On Error Resume Next

'    '使用 Is Nothing 方法判斷指定名字的模塊是否已經存在，注意需要寫上語句 On Error Resume Next，不然當沒有找到指定的模塊時會報錯
'    If ThisWorkbook.VBProject.VBComponents(moduleName) Is Nothing Then

'        Debug.Print "Module ( " & moduleName & " ) is Nothing."

'    Else

'        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(moduleName)  '移除模塊

'        '模塊移除之後，再次使用 Is Nothing 方法判斷指定名字的模塊是否已經存在，注意需要寫上語句 On Error Resume Next，不然當沒有找到指定的模塊時會報錯
'        If ThisWorkbook.VBProject.VBComponents(moduleName) Is Nothing Then

'            'Debug.Print "Module ( " & moduleName & " ) be removed."

'        Else

'            Debug.Print "remove Module ( " & moduleName & " ) error, Module ( " & moduleName & " ) cannot be remove."
'            'Exit Sub

'        End If

'    End If

'End Sub

