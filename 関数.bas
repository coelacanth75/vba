Attribute VB_Name = "関数"
'----------------------------------------------------------------------------------------------
' auther : R.Sugita
' proc   : name                   type description
'          ---------------------- ---- --------------------------------------------------------
'          reduction              sub  選択しているセルに「縮小して全体を表示」を設定
'          myJoin                 func 指定した範囲のデータを結合
'          insertJoin             func INSERT文の生成
'          snakeToCamel           func スネークケースをキャメルケースに変換
'          countIfVBA             func 検索範囲内に検索値のデータが何行存在するかカウント
'          countRangeVlookup      func 検索範囲内に検索値範囲のデータが何行存在するかカウント
'          countNonOverlap        func 検索範囲内の、重複していない行数を取得
'          dispNonOverlap         sub  検索範囲内の、重複していないデータを取得
'          deleteCanceledChar     sub  選択範囲の取り消し線のついた文字を削除
'          delete_style           sub  アクティブシートのスタイル削除
'          lenb_utf8              func UTF-8のバイト数を取得
'          getEnableStr           func 有効な文字（取り消しされてない）のみ取得
'          searchRev              func 検索対象内の、検索文字の最後の一致のインデックスを返す
'          sheetAllVisible        sub  ブックのシート全表示
'          createSheetTitleList   sub  目次シートの作成
'          countStartString       func 対象文字列の、先頭何文字が検索文字列と一致するかを返す
'          tree                   sub  親子関係のデータをツリー表示する
'          searchChild            func プロシージャ「tree」の内部関数
'          changeCharColor        sub  特定の文字列の色を変える
'          swichDisplayPageBreaks sub  改ページの表示の切り替え
'          setDataEvidenceLayout  sub  罫線、ヘッダ色の設定
'----------------------------------------------------------------------------------------------


' 指定した範囲のデータを結合
Function myJoin(範囲 As Range, Optional 区切り文字 As String) As Variant
Dim c As Range, buf As String
   If 範囲.Rows.count = 1 Or 範囲.Columns.count = 1 Then
      For Each c In 範囲
         If c.Value <> "" Then
           buf = buf & 区切り文字 & c.Value
         End If
      Next c
      If 区切り文字 <> "" Then
         myJoin = Mid$(buf, Len(区切り文字) + 1)
         Else
         myJoin = buf
      End If
   Else
      myJoin = CVErr(xlErrRef)  'エラー値
   End If
End Function

' INSERT文の生成
' 接頭辞列、接尾辞列、カラム名範囲列、データ範囲がある状態を想定。
Function insertJoin(データ範囲 As Range, カラム名範囲 As Range, Optional テーブル名 As String, Optional 接頭辞 As Range, Optional 接尾辞 As Range) As Variant
    Dim i As Long
    Dim columnNames As String
    Dim dataValues  As String
    Dim temp        As String
    Dim dataArray   As Variant
    Dim columnArray As Variant
    Dim prefixArray As Variant
    Dim suffixArray As Variant
    Dim result      As String
    
    ' 変数の初期化
    dataArray = データ範囲
    columnArray = カラム名範囲
    
    If Not 接頭辞 Is Nothing Then
      prefixArray = 接頭辞
    End If
    
    If Not 接尾辞 Is Nothing Then
      suffixArray = 接尾辞
    End If
    
    If テーブル名 = "" Then
      columnNames = "( "
    Else
      columnNames = "INSERT INTO " & テーブル名 & " ( "
    End If
    
    dataValues = "VALUES ( "
    
    ' 処理開始
    If データ範囲.Rows.count = 1 And カラム名範囲.Rows.count = 1 Then
    
       For i = 1 To データ範囲.count
       
          ' 一時変数の初期化
          temp = ""
       
          If dataArray(1, i) <> "" Then
            columnNames = columnNames & columnArray(1, i) & " , "
            
            If Not 接頭辞 Is Nothing Then
              If prefixArray(1, i) <> "" Then
                temp = prefixArray(1, i)
              End If
            End If
            
            temp = temp & dataArray(1, i)
            
            If Not 接尾辞 Is Nothing Then
              If suffixArray(1, i) <> "" Then
                temp = temp & suffixArray(1, i)
              End If
            End If
            
            dataValues = dataValues & temp & " , "
          End If
       Next
       
       result = Mid(columnNames, 1, Len(columnNames) - 2) & ") "
       result = result & Mid(dataValues, 1, Len(dataValues) - 2) & ");"
       
       insertJoin = result
       
    Else
       insertJoin = CVErr(xlErrRef)  'エラー値
    End If
End Function

' スネークケースをキャメルケースに変換
Function snakeToCamel(対象 As String) As String
Dim i As Integer, buf As String
   If 対象 <> "" Then
   
      buf = 対象
      
      For i = 1 To 26
        buf = Replace(buf, "_" & Chr(i + 96), Chr(i + 64))
      Next
      
      snakeToCamel = buf
      
   Else
   
      snakeToCamel = ""
      
   End If
End Function

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

' 検索範囲内の、重複していない行数を取得
Function countNonOverlap(検索範囲 As Range) As Long
    Dim countCount As Long
    Dim rowCount As Long
    Dim matchCount As Long
    Dim dictionary As Scripting.dictionary
    Dim text As String
    Dim i As Long
    Dim j As Long
    
    columnCount = 検索範囲.Columns.count
    rowCount = 検索範囲.Rows.count
    Set dictionary = New Scripting.dictionary
    
    matchCount = 0
    
    For i = 1 To rowCount
        For j = 1 To columnCount
            text = 検索範囲(i, j).Value
            '1カラムづつ比較し、一致しない場合はカウントアップ
            If dictionary.Exists(text) <> True Then
                dictionary.Add text, text
                matchCount = matchCount + 1
            End If
        Next
    Next
    
    countNonOverlap = matchCount
    
End Function

' 検索範囲内の、重複していないデータを取得
Sub dispNonOverlap()
    Dim columnOrigin As Long
    Dim rowOrigin As Long
    Dim columnCount As Long
    Dim rowCount As Long
    Dim matchCount As Long
    Dim dictionary As Scripting.dictionary
    Dim text As String
    Dim i As Long
    Dim j As Long
    
    columnOrigin = Selection(1, 1).Column - 1
    rowOrigin = Selection(1, 1).row - 1
    columnCount = Selection.Columns.count
    rowCount = Selection.Rows.count
    Set dictionary = New Scripting.dictionary
    
    matchCount = 0
    
    For i = 1 To rowCount
        If Rows(rowOrigin + i).Hidden <> True Then
        For j = 1 To columnCount
            If Columns(columnOrigin + j).Hidden <> True Then
                text = Selection(i, j).Value
                '1カラムづつ比較し、一致しない場合は出力
                If dictionary.Exists(text) <> True Then
                    dictionary.Add text, text
                    Debug.Print text
                End If
            End If
        Next
        End If
    Next
    
End Sub

' 選択範囲の取り消し線のついた文字を削除
Sub deleteCanceledChar()
    Dim textBefore
    Dim textAfter

    For Each myCell In Selection
        textBefore = myCell.Value
        textAfter = ""
        If textBefore <> "" Then
            For i = 1 To Len(textBefore)
                If myCell.Characters(start:=i, length:=1).Font.Strikethrough = False Then
                    textAfter = textAfter & Mid(textBefore, i, 1)
                End If
            Next i
            myCell.Value = textAfter
        End If
    Next
End Sub

' アクティブシートのスタイル削除
Sub delete_style()
    Dim m()
    Dim i
    Dim j
    j = ActiveWorkbook.Styles.count
    ReDim m(j)
    For i = 1 To j
        m(i) = ActiveWorkbook.Styles(i).Name
    Next
    For i = 1 To j
        If InStr("Hyperlink,Normal,Followed Hyperlink", m(i)) = 0 Then
            ActiveWorkbook.Styles(m(i)).Delete
        End If
    Next

End Sub

' UTF-8のバイト数を取得
Function lenb_utf8(対象 As String)
Dim UTF8 As Object
Dim target As String

target = 対象

On Error GoTo errh

Set UTF8 = CreateObject("System.Text.UTF8Encoding")
lenb_utf8 = UTF8.GetByteCount_2(target)


errh:
If Err.Number <> 0 Then
lenb_utf8 = CVErr(xlErrRef)  'エラー値
End If

Set UTF8 = Nothing
End Function

' 有効な文字（取り消しされてない）のみ取得　※性能良くない
Function getEnableStr(cell As Range, Optional start As Integer = 1, Optional length As Integer = 0)
    Dim char As Characters
    Dim half As Integer
    Dim result As String
    
    ' 開始位置引数が無効の場合、1とする。
    If start <= 0 Then
        start = 1
    End If
    
    ' 長さ引数が無効の場合、開始位置以降すべてとする。
    If length <= 0 Then
        length = Len(cell.text) - start + 1
    End If
    
    ' 二分木の要領で処理
    Set char = cell.Characters(start, length)
    Select Case char.Font.Strikethrough
    Case False
        ' すべての文字が通常文字
        result = char.text
    Case True
        ' すべての文字が取消文字
        result = ""
    Case Else
        ' 取消文字、通常文字の混在
        half = length / 2
        result = getEnableStr(cell, start, half)
        result = result + getEnableStr(cell, start + half, length - half)
    End Select
    
    getEnableStr = result
End Function

' 検索対象内の、検索文字の最後の一致のインデックスを返す
Function searchRev(検索文字 As String, 検索対象 As String)
    searchRev = InStrRev(検索対象, 検索文字)
End Function

' ブックのシート全表示
Sub sheetAllVisible()
    Dim sh As Object
    
    For Each sh In Sheets
        sh.Visible = True
    Next sh
End Sub

' 目次シートの作成
Sub createSheetTitleList()
    Dim i As Long
    Worksheets.Add before:=Worksheets(1)
    ActiveSheet.Name = "目次"
    
    For i = 1 To Sheets.count
    If Worksheets(i).Name <> "目次" Then
        Range("B" & (i + 3)).Value = Worksheets(i).Name
            Worksheets(1).Hyperlinks.Add _
              Anchor:=Range("B" & (i + 3)), _
              Address:="", _
              SubAddress:="'" & Worksheets(i).Name & "'" & "!A1", _
              TextToDisplay:=Worksheets(i).Name
    End If
    Next i

End Sub

' target（第一引数）で指定した文字列の、先頭何文字がsearchStr（第二引数）と一致するかを返す
Function countStartString(target As String, searchStr As String) As Long
    Dim count
    Dim i
    
    count = 0
    
    For i = 1 To Len(target)
        If Mid(target, i, 1) = searchStr Then
            count = count + 1
        Else
            Exit For
        End If
    Next i
    
    countStartString = count

End Function

' 親子関係のデータをツリー表示する
' 変数targetは1列目にキー値、2列目にレベルが格納されるものとする。
Sub tree()
    Dim arrayList() As Variant
    Dim i As Integer
    Dim j As Integer

    ' 変数arrayListには1行目に親のキー値、2行目に子のキー値を格納する。
    ReDim arrayList(Selection.count / 2, 2)
    
    For i = 1 To Selection.count / 2
    
        ' レベルが0の場合、親は無し
        If Selection(i, 2) = 0 Then
            arrayList(i, 1) = ""
        End If
        
        If arrayList(i, 0) = "" Then
            ' 子を検索する
            arrayList = searchChild(Selection, arrayList, i)
        Else
            ' 処理済みの場合、スキップする
            'Debug.Print Selection(i, 1) & "は処理済みのため、スキップ (i=" & i & ")"
        End If
        
        ' 出力する
        Debug.Print Selection(i, 1) & vbTab & Selection(i, 2) & vbTab & arrayList(i, 1) & vbTab & arrayList(i, 2)
        
    Next i
    

End Sub

' プロシージャ「tree」の内部関数
Function searchChild(target As Range, arrayList As Variant, i As Integer) As Variant
    Dim arrayChild As Variant
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    
    ' 子を検索する
    For j = i + 1 To target.count / 2
       
        If target(i, 2) + 1 = target(j, 2) And arrayList(j, 0) = "" Then
            ' 子の場合
            'Debug.Print target(j, 1) & "は" & target(i, 1) & "の子である (i=" & i & ", j=" & j & ")"
            
            If arrayList(i, 2) <> "" Then
                arrayList(i, 2) = arrayList(i, 2) & ","
            End If
            
            arrayList(j, 1) = target(i, 1)
            arrayList(i, 2) = arrayList(i, 2) & target(j, 1)
            
        End If
        
    Next j
        
    If arrayList(i, 2) <> "" Then
       ' 子が存在する場合
       
        ' 子を配列にする。
        arrayChild = Split(arrayList(i, 2), ",")
        
        For k = UBound(arrayChild) To 0 Step -1
            l = target.Find(arrayChild(k)).row - 233
            
            If arrayList(l, 0) = "" Then
                arrayList = searchChild(target, arrayList, l)
            Else
                ' 処理済みの場合、スキップする
                'Debug.Print target(l, 1) & "は処理済みのため、スキップ (i=" & i & ", k=" & k & ", l=" & l & ")"
            End If
Continue:
        Next k
        
    End If
    
    
    arrayList(i, 0) = "Y"
    searchChild = arrayList

End Function

' 特定の文字列の色を変える
Sub changeCharColor()
    Dim rng As Range
    Dim ptr As Integer
    Const tStr As String = "AP_INVOICES" 'ここに色を変える文字列を書く
    For Each rng In ActiveSheet.Cells.SpecialCells(xlCellTypeConstants, xlTextValues)
        ptr = InStr(rng.Value, tStr)
        If ptr > 0 Then
            rng.Characters(start:=ptr, length:=Len(tStr)).Font.ColorIndex = 3
        End If
    Next rng
End Sub

' 改ページの表示の切り替え
Sub swichDisplayPageBreaks()
    If ActiveSheet.DisplayPageBreaks Then
        ActiveSheet.DisplayPageBreaks = False
    Else
        ActiveSheet.DisplayPageBreaks = True
    End If
End Sub

' データエビデンス用（罫線、ヘッダ色）
Sub setDataEvidenceLayout()
    Selection.Borders.LineStyle = True
    
    Dim row
    Dim col
    row = Selection(1).row
    col = Selection(Selection.count).Column
    
    With Range(Selection(1), Cells(row, col)).Interior
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
    End With
End Sub

' 文字列検索
Function findVBA(検索文字列 As String, 対象 As String, 要素数 As Long) As Long
    On Error GoTo ErrorHandler
    findVBA = WorksheetFunction.Find(Chr(16), WorksheetFunction.Substitute(対象, 検索文字列, Chr(16), 要素数))
    Exit Function
ErrorHandler:
    findVBA = 0
End Function

' 文字列検索
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








