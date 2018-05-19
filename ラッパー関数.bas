Attribute VB_Name = "���b�p�[�֐�"
'----------------------------------------------------------------------------------------------
' auther : R.Sugita
' proc   : name                   type description
'          ---------------------- ---- --------------------------------------------------------
'          countIfVBA             func �����͈͓��Ɍ����l�̃f�[�^�����s���݂��邩�J�E���g
'          countRangeVlookup      func �����͈͓��Ɍ����l�͈͂̃f�[�^�����s���݂��邩�J�E���g
'          searchRev              func �����Ώۓ��́A���������̍Ō�̈�v�̃C���f�b�N�X��Ԃ�
'          findVBA                func �����񌟍�
'          midVBA                 func ������擾
'----------------------------------------------------------------------------------------------


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

' �����Ώۓ��́A���������̍Ō�̈�v�̃C���f�b�N�X��Ԃ�
Function searchRev(�������� As String, �����Ώ� As String)
    searchRev = InStrRev(�����Ώ�, ��������)
End Function

' �����񌟍�
Function findVBA(���������� As String, �Ώ� As String, �v�f�� As Long) As Long
    On Error GoTo ErrorHandler
    findVBA = WorksheetFunction.Find(Chr(16), WorksheetFunction.Substitute(�Ώ�, ����������, Chr(16), �v�f��))
    Exit Function
ErrorHandler:
    findVBA = 0
End Function

' ������擾
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










