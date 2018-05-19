Attribute VB_Name = "ラッパー関数"
'----------------------------------------------------------------------------------------------
' auther : R.Sugita
' proc   : name                   type description
'          ---------------------- ---- --------------------------------------------------------
'          countIfVBA             func 検索範囲内に検索値のデータが何行存在するかカウント
'          countRangeVlookup      func 検索範囲内に検索値範囲のデータが何行存在するかカウント
'          searchRev              func 検索対象内の、検索文字の最後の一致のインデックスを返す
'          findVBA                func 文字列検索
'          midVBA                 func 文字列取得
'----------------------------------------------------------------------------------------------


' 検索範囲内に検索値のデータが何行存在するかカウント（COUNTIFで文字数エラーとなった時用）
Function countIfVBA(検索値 As String, 検索範囲 As Range) As Long
    Dim rowCount As Long
    Dim matchCount As Long
    Dim searchArray As Variant
    Dim i As Long
    Dim isMatch As Boolean
    
    If 検索範囲.Columns.count <> 1 Then
        countIfVBA = vbError
    End If
    
    rowCount = 検索範囲.Rows.count
    searchArray = 検索範囲
    
    matchCount = 0
    isMatch = True
    
    For i = 1 To rowCount
        '1カラムづつ比較し、一致しない場合はカウントアップして処理を抜ける
        If 検索値 <> searchArray(i, 1) Then
            isMatch = False
            Exit For
        End If
            
        If isMatch Then
            matchCount = matchCount + 1
        Else
            isMatch = True
        End If
    Next
    
    countIfVBA = matchCount
    
End Function

' 検索範囲内に検索値範囲のデータが何行存在するかカウント
Function countRangeVlookup(検索値範囲 As Range, 検索範囲 As Range) As Long
    Dim columnCount As Long
    Dim rowCount As Long
    Dim matchCount As Long
    Dim targetArray As Variant
    Dim searchArray As Variant
    Dim i As Long
    Dim j As Long
    Dim isMatch As Boolean
    
    columnCount = 検索値範囲.Columns.count
    rowCount = 検索範囲.Rows.count
    targetArray = 検索値範囲
    searchArray = 検索範囲
    
    If columnCount <> 検索範囲.Columns.count _
        Or 検索値範囲.Rows.count > 1 Then
        countRangeVlookup = vbError
    End If
    
    matchCount = 0
    isMatch = True
    
    For i = 1 To rowCount
        For j = 1 To columnCount
            '1カラムづつ比較し、一致しない場合はカウントアップして処理を抜ける
            If targetArray(1, j) <> searchArray(i, j) Then
                isMatch = False
                Exit For
            End If
        Next
            
        If isMatch Then
            matchCount = matchCount + 1
        Else
            isMatch = True
        End If
    Next
    
    countRangeVlookup = matchCount
    
End Function

' 検索対象内の、検索文字の最後の一致のインデックスを返す
Function searchRev(検索文字 As String, 検索対象 As String)
    searchRev = InStrRev(検索対象, 検索文字)
End Function

' 文字列検索
Function findVBA(検索文字列 As String, 対象 As String, 要素数 As Long) As Long
    On Error GoTo ErrorHandler
    findVBA = WorksheetFunction.Find(Chr(16), WorksheetFunction.Substitute(対象, 検索文字列, Chr(16), 要素数))
    Exit Function
ErrorHandler:
    findVBA = 0
End Function

' 文字列取得
Function midVBA(対象 As String, 開始文字列 As String, 終了文字列 As String, 要素数 As Long, Optional flag As Boolean = True) As String
    Dim startStrIndex
    startStrIndex = findVBA(開始文字列, 対象, 要素数)
    
    If startStrIndex = 0 Then
        GoTo ErrorHandler
    End If
    
    Dim endStrIndex
    endStrIndex = InStr(startStrIndex, 対象, 終了文字列)
    
    If flag Then
        startStrIndex = startStrIndex + Len(開始文字列)
        endStrIndex = endStrIndex - Len(終了文字列)
    End If
    
    On Error GoTo ErrorHandler
    midVBA = Mid(対象, startStrIndex, endStrIndex - startStrIndex + 1)
    Exit Function
ErrorHandler:
    midVBA = ""
End Function










