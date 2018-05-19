Attribute VB_Name = "�֐�"
'----------------------------------------------------------------------------------------------
' auther : R.Sugita
' proc   : name                   type description
'          ---------------------- ---- --------------------------------------------------------
'          reduction              sub  �I�����Ă���Z���Ɂu�k�����đS�̂�\���v��ݒ�
'          myJoin                 func �w�肵���͈͂̃f�[�^������
'          insertJoin             func INSERT���̐���
'          snakeToCamel           func �X�l�[�N�P�[�X���L�������P�[�X�ɕϊ�
'          countIfVBA             func �����͈͓��Ɍ����l�̃f�[�^�����s���݂��邩�J�E���g
'          countRangeVlookup      func �����͈͓��Ɍ����l�͈͂̃f�[�^�����s���݂��邩�J�E���g
'          countNonOverlap        func �����͈͓��́A�d�����Ă��Ȃ��s�����擾
'          dispNonOverlap         sub  �����͈͓��́A�d�����Ă��Ȃ��f�[�^���擾
'          deleteCanceledChar     sub  �I��͈͂̎��������̂����������폜
'          delete_style           sub  �A�N�e�B�u�V�[�g�̃X�^�C���폜
'          lenb_utf8              func UTF-8�̃o�C�g�����擾
'          getEnableStr           func �L���ȕ����i����������ĂȂ��j�̂ݎ擾
'          searchRev              func �����Ώۓ��́A���������̍Ō�̈�v�̃C���f�b�N�X��Ԃ�
'          sheetAllVisible        sub  �u�b�N�̃V�[�g�S�\��
'          createSheetTitleList   sub  �ڎ��V�[�g�̍쐬
'          countStartString       func �Ώە�����́A�擪������������������ƈ�v���邩��Ԃ�
'          tree                   sub  �e�q�֌W�̃f�[�^���c���[�\������
'          searchChild            func �v���V�[�W���utree�v�̓����֐�
'          changeCharColor        sub  ����̕�����̐F��ς���
'          swichDisplayPageBreaks sub  ���y�[�W�̕\���̐؂�ւ�
'          setDataEvidenceLayout  sub  �r���A�w�b�_�F�̐ݒ�
'----------------------------------------------------------------------------------------------


' �w�肵���͈͂̃f�[�^������
Function myJoin(�͈� As Range, Optional ��؂蕶�� As String) As Variant
Dim c As Range, buf As String
   If �͈�.Rows.count = 1 Or �͈�.Columns.count = 1 Then
      For Each c In �͈�
         If c.Value <> "" Then
           buf = buf & ��؂蕶�� & c.Value
         End If
      Next c
      If ��؂蕶�� <> "" Then
         myJoin = Mid$(buf, Len(��؂蕶��) + 1)
         Else
         myJoin = buf
      End If
   Else
      myJoin = CVErr(xlErrRef)  '�G���[�l
   End If
End Function

' INSERT���̐���
' �ړ�����A�ڔ�����A�J�������͈͗�A�f�[�^�͈͂������Ԃ�z��B
Function insertJoin(�f�[�^�͈� As Range, �J�������͈� As Range, Optional �e�[�u���� As String, Optional �ړ��� As Range, Optional �ڔ��� As Range) As Variant
    Dim i As Long
    Dim columnNames As String
    Dim dataValues  As String
    Dim temp        As String
    Dim dataArray   As Variant
    Dim columnArray As Variant
    Dim prefixArray As Variant
    Dim suffixArray As Variant
    Dim result      As String
    
    ' �ϐ��̏�����
    dataArray = �f�[�^�͈�
    columnArray = �J�������͈�
    
    If Not �ړ��� Is Nothing Then
      prefixArray = �ړ���
    End If
    
    If Not �ڔ��� Is Nothing Then
      suffixArray = �ڔ���
    End If
    
    If �e�[�u���� = "" Then
      columnNames = "( "
    Else
      columnNames = "INSERT INTO " & �e�[�u���� & " ( "
    End If
    
    dataValues = "VALUES ( "
    
    ' �����J�n
    If �f�[�^�͈�.Rows.count = 1 And �J�������͈�.Rows.count = 1 Then
    
       For i = 1 To �f�[�^�͈�.count
       
          ' �ꎞ�ϐ��̏�����
          temp = ""
       
          If dataArray(1, i) <> "" Then
            columnNames = columnNames & columnArray(1, i) & " , "
            
            If Not �ړ��� Is Nothing Then
              If prefixArray(1, i) <> "" Then
                temp = prefixArray(1, i)
              End If
            End If
            
            temp = temp & dataArray(1, i)
            
            If Not �ڔ��� Is Nothing Then
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
       insertJoin = CVErr(xlErrRef)  '�G���[�l
    End If
End Function

' �X�l�[�N�P�[�X���L�������P�[�X�ɕϊ�
Function snakeToCamel(�Ώ� As String) As String
Dim i As Integer, buf As String
   If �Ώ� <> "" Then
   
      buf = �Ώ�
      
      For i = 1 To 26
        buf = Replace(buf, "_" & Chr(i + 96), Chr(i + 64))
      Next
      
      snakeToCamel = buf
      
   Else
   
      snakeToCamel = ""
      
   End If
End Function

' �����͈͓��Ɍ����l�̃f�[�^�����s���݂��邩�J�E���g�iCOUNTIF�ŕ������G���[�ƂȂ������p�j
Function countIfVBA(�����l As String, �����͈� As Range) As Long
    Dim rowCount As Long
    Dim matchCount As Long
    Dim searchArray As Variant
    Dim i As Long
    Dim isMatch As Boolean
    
    If �����͈�.Columns.count <> 1 Then
        countIfVBA = vbError
    End If
    
    rowCount = �����͈�.Rows.count
    searchArray = �����͈�
    
    matchCount = 0
    isMatch = True
    
    For i = 1 To rowCount
        '1�J�����Â�r���A��v���Ȃ��ꍇ�̓J�E���g�A�b�v���ď����𔲂���
        If �����l <> searchArray(i, 1) Then
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

' �����͈͓��Ɍ����l�͈͂̃f�[�^�����s���݂��邩�J�E���g
Function countRangeVlookup(�����l�͈� As Range, �����͈� As Range) As Long
    Dim columnCount As Long
    Dim rowCount As Long
    Dim matchCount As Long
    Dim targetArray As Variant
    Dim searchArray As Variant
    Dim i As Long
    Dim j As Long
    Dim isMatch As Boolean
    
    columnCount = �����l�͈�.Columns.count
    rowCount = �����͈�.Rows.count
    targetArray = �����l�͈�
    searchArray = �����͈�
    
    If columnCount <> �����͈�.Columns.count _
        Or �����l�͈�.Rows.count > 1 Then
        countRangeVlookup = vbError
    End If
    
    matchCount = 0
    isMatch = True
    
    For i = 1 To rowCount
        For j = 1 To columnCount
            '1�J�����Â�r���A��v���Ȃ��ꍇ�̓J�E���g�A�b�v���ď����𔲂���
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

' �����͈͓��́A�d�����Ă��Ȃ��s�����擾
Function countNonOverlap(�����͈� As Range) As Long
    Dim countCount As Long
    Dim rowCount As Long
    Dim matchCount As Long
    Dim dictionary As Scripting.dictionary
    Dim text As String
    Dim i As Long
    Dim j As Long
    
    columnCount = �����͈�.Columns.count
    rowCount = �����͈�.Rows.count
    Set dictionary = New Scripting.dictionary
    
    matchCount = 0
    
    For i = 1 To rowCount
        For j = 1 To columnCount
            text = �����͈�(i, j).Value
            '1�J�����Â�r���A��v���Ȃ��ꍇ�̓J�E���g�A�b�v
            If dictionary.Exists(text) <> True Then
                dictionary.Add text, text
                matchCount = matchCount + 1
            End If
        Next
    Next
    
    countNonOverlap = matchCount
    
End Function

' �����͈͓��́A�d�����Ă��Ȃ��f�[�^���擾
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
                '1�J�����Â�r���A��v���Ȃ��ꍇ�͏o��
                If dictionary.Exists(text) <> True Then
                    dictionary.Add text, text
                    Debug.Print text
                End If
            End If
        Next
        End If
    Next
    
End Sub

' �I��͈͂̎��������̂����������폜
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

' �A�N�e�B�u�V�[�g�̃X�^�C���폜
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

' UTF-8�̃o�C�g�����擾
Function lenb_utf8(�Ώ� As String)
Dim UTF8 As Object
Dim target As String

target = �Ώ�

On Error GoTo errh

Set UTF8 = CreateObject("System.Text.UTF8Encoding")
lenb_utf8 = UTF8.GetByteCount_2(target)


errh:
If Err.Number <> 0 Then
lenb_utf8 = CVErr(xlErrRef)  '�G���[�l
End If

Set UTF8 = Nothing
End Function

' �L���ȕ����i����������ĂȂ��j�̂ݎ擾�@�����\�ǂ��Ȃ�
Function getEnableStr(cell As Range, Optional start As Integer = 1, Optional length As Integer = 0)
    Dim char As Characters
    Dim half As Integer
    Dim result As String
    
    ' �J�n�ʒu�����������̏ꍇ�A1�Ƃ���B
    If start <= 0 Then
        start = 1
    End If
    
    ' ���������������̏ꍇ�A�J�n�ʒu�ȍ~���ׂĂƂ���B
    If length <= 0 Then
        length = Len(cell.text) - start + 1
    End If
    
    ' �񕪖؂̗v�̂ŏ���
    Set char = cell.Characters(start, length)
    Select Case char.Font.Strikethrough
    Case False
        ' ���ׂĂ̕������ʏ핶��
        result = char.text
    Case True
        ' ���ׂĂ̕������������
        result = ""
    Case Else
        ' ��������A�ʏ핶���̍���
        half = length / 2
        result = getEnableStr(cell, start, half)
        result = result + getEnableStr(cell, start + half, length - half)
    End Select
    
    getEnableStr = result
End Function

' �����Ώۓ��́A���������̍Ō�̈�v�̃C���f�b�N�X��Ԃ�
Function searchRev(�������� As String, �����Ώ� As String)
    searchRev = InStrRev(�����Ώ�, ��������)
End Function

' �u�b�N�̃V�[�g�S�\��
Sub sheetAllVisible()
    Dim sh As Object
    
    For Each sh In Sheets
        sh.Visible = True
    Next sh
End Sub

' �ڎ��V�[�g�̍쐬
Sub createSheetTitleList()
    Dim i As Long
    Worksheets.Add before:=Worksheets(1)
    ActiveSheet.Name = "�ڎ�"
    
    For i = 1 To Sheets.count
    If Worksheets(i).Name <> "�ڎ�" Then
        Range("B" & (i + 3)).Value = Worksheets(i).Name
            Worksheets(1).Hyperlinks.Add _
              Anchor:=Range("B" & (i + 3)), _
              Address:="", _
              SubAddress:="'" & Worksheets(i).Name & "'" & "!A1", _
              TextToDisplay:=Worksheets(i).Name
    End If
    Next i

End Sub

' target�i�������j�Ŏw�肵��������́A�擪��������searchStr�i�������j�ƈ�v���邩��Ԃ�
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

' �e�q�֌W�̃f�[�^���c���[�\������
' �ϐ�target��1��ڂɃL�[�l�A2��ڂɃ��x�����i�[�������̂Ƃ���B
Sub tree()
    Dim arrayList() As Variant
    Dim i As Integer
    Dim j As Integer

    ' �ϐ�arrayList�ɂ�1�s�ڂɐe�̃L�[�l�A2�s�ڂɎq�̃L�[�l���i�[����B
    ReDim arrayList(Selection.count / 2, 2)
    
    For i = 1 To Selection.count / 2
    
        ' ���x����0�̏ꍇ�A�e�͖���
        If Selection(i, 2) = 0 Then
            arrayList(i, 1) = ""
        End If
        
        If arrayList(i, 0) = "" Then
            ' �q����������
            arrayList = searchChild(Selection, arrayList, i)
        Else
            ' �����ς݂̏ꍇ�A�X�L�b�v����
            'Debug.Print Selection(i, 1) & "�͏����ς݂̂��߁A�X�L�b�v (i=" & i & ")"
        End If
        
        ' �o�͂���
        Debug.Print Selection(i, 1) & vbTab & Selection(i, 2) & vbTab & arrayList(i, 1) & vbTab & arrayList(i, 2)
        
    Next i
    

End Sub

' �v���V�[�W���utree�v�̓����֐�
Function searchChild(target As Range, arrayList As Variant, i As Integer) As Variant
    Dim arrayChild As Variant
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    
    ' �q����������
    For j = i + 1 To target.count / 2
       
        If target(i, 2) + 1 = target(j, 2) And arrayList(j, 0) = "" Then
            ' �q�̏ꍇ
            'Debug.Print target(j, 1) & "��" & target(i, 1) & "�̎q�ł��� (i=" & i & ", j=" & j & ")"
            
            If arrayList(i, 2) <> "" Then
                arrayList(i, 2) = arrayList(i, 2) & ","
            End If
            
            arrayList(j, 1) = target(i, 1)
            arrayList(i, 2) = arrayList(i, 2) & target(j, 1)
            
        End If
        
    Next j
        
    If arrayList(i, 2) <> "" Then
       ' �q�����݂���ꍇ
       
        ' �q��z��ɂ���B
        arrayChild = Split(arrayList(i, 2), ",")
        
        For k = UBound(arrayChild) To 0 Step -1
            l = target.Find(arrayChild(k)).row - 233
            
            If arrayList(l, 0) = "" Then
                arrayList = searchChild(target, arrayList, l)
            Else
                ' �����ς݂̏ꍇ�A�X�L�b�v����
                'Debug.Print target(l, 1) & "�͏����ς݂̂��߁A�X�L�b�v (i=" & i & ", k=" & k & ", l=" & l & ")"
            End If
Continue:
        Next k
        
    End If
    
    
    arrayList(i, 0) = "Y"
    searchChild = arrayList

End Function

' ����̕�����̐F��ς���
Sub changeCharColor()
    Dim rng As Range
    Dim ptr As Integer
    Const tStr As String = "AP_INVOICES" '�����ɐF��ς��镶���������
    For Each rng In ActiveSheet.Cells.SpecialCells(xlCellTypeConstants, xlTextValues)
        ptr = InStr(rng.Value, tStr)
        If ptr > 0 Then
            rng.Characters(start:=ptr, length:=Len(tStr)).Font.ColorIndex = 3
        End If
    Next rng
End Sub

' ���y�[�W�̕\���̐؂�ւ�
Sub swichDisplayPageBreaks()
    If ActiveSheet.DisplayPageBreaks Then
        ActiveSheet.DisplayPageBreaks = False
    Else
        ActiveSheet.DisplayPageBreaks = True
    End If
End Sub

' �f�[�^�G�r�f���X�p�i�r���A�w�b�_�F�j
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

' �����񌟍�
Function findVBA(���������� As String, �Ώ� As String, �v�f�� As Long) As Long
    On Error GoTo ErrorHandler
    findVBA = WorksheetFunction.Find(Chr(16), WorksheetFunction.Substitute(�Ώ�, ����������, Chr(16), �v�f��))
    Exit Function
ErrorHandler:
    findVBA = 0
End Function

' �����񌟍�
Function midVBA(�Ώ� As String, �J�n������ As String, �I�������� As String, �v�f�� As Long, Optional flag As Boolean = True) As String
    Dim startStrIndex
    startStrIndex = findVBA(�J�n������, �Ώ�, �v�f��)
    
    If startStrIndex = 0 Then
        GoTo ErrorHandler
    End If
    
    Dim endStrIndex
    endStrIndex = InStr(startStrIndex, �Ώ�, �I��������)
    
    If flag Then
        startStrIndex = startStrIndex + Len(�J�n������)
        endStrIndex = endStrIndex - Len(�I��������)
    End If
    
    On Error GoTo ErrorHandler
    midVBA = Mid(�Ώ�, startStrIndex, endStrIndex - startStrIndex + 1)
    Exit Function
ErrorHandler:
    midVBA = ""
End Function








