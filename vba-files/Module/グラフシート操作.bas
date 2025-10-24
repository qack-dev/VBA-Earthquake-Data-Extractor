Option Explicit

' key���k���n�Avalue���񐔂�Dictionary���쐬���o��
Public Sub makeLocateDict()
    ' �ϐ��錾
    Dim dict As Object
    Dim locateRange As Range
    Dim r As Range
    Dim key As Variant
    Dim tmpRow As Long
    ' ���
    Set dict = CreateObject("Scripting.Dictionary")
    tmpRow = 2
    With EXTRACT_SHEET
        Set locateRange = .Range( _
            .Cells(3, locateCol), _
            .Cells(.Rows.Count, locateCol).End(xlUp) _
        )
    End With
    ' �k���n��������[�v���]�L
    For Each r In locateRange
        ' ���Ɏ�����key�ɑ��݂�����
        If dict.exists(r.Value) Then
            dict(r.Value) = dict(r.Value) + 1
        ' ���݂��Ȃ�������
        Else
            dict(r.Value) = 1
        End If
    Next r
    ' �o��
    With GRAPH_SHEET
        ' �\��������
        .Activate
        .Range(Columns(1), Columns(2)).Delete
        ' ���o���s����
        .Cells(1, 1).Value = "�k���n��"
        .Cells(1, 2).Value = "������"
        For Each key In dict.Keys
            ' �\�̃��R�[�h����
            .Cells(tmpRow, 1).Value = key
            .Cells(tmpRow, 2).Value = dict(key)
            tmpRow = tmpRow + 1
        Next key
        ' ���`
        With .Range(.Cells(1, 1), .Cells(tmpRow - 1, 2))
            .EntireColumn.AutoFit
            .Borders.LineStyle = xlContinuous
        End With
    End With
    Cells(2, 1).Select
End Sub

' �\�[�g
Public Sub sortTableExcel2007()
    Dim bodyRange As Range
    With GRAPH_SHEET
        Set bodyRange = .Range( _
            .Cells(2, 1), _
            .Cells(.Rows.Count, 2).End(xlUp) _
        )
    End With
    With GRAPH_SHEET.Sort
        ' ���݂̕��ёւ����N���A
        .SortFields.Clear
        ' �����񐔂ō~��
        .SortFields.Add _
            key:=Cells(2, 2), _
            Order:=xlDescending
        .SetRange bodyRange
        .Header = xlNo
        .Orientation = xlTopToBottom
        ' �k���n���ŏ���
        .SortFields.Add _
            key:=Cells(2, 1), _
            Order:=xlAscending
        .SetRange bodyRange
        .Header = xlNo
        .Orientation = xlTopToBottom
        ' �K�p
        .Apply
    End With
End Sub

' �����񐔂�20�ʂ܂łŏc�_�O���t�쐬
Public Sub makeGraph()
    ' �ϐ��錾�E���
    Dim targetRange As Range
    Set targetRange _
    = ActiveSheet.Range(Cells(3, 4), Cells(30, 10))
    With GRAPH_SHEET
        .Activate
        
        ' �G���[���������Ă������𑱂���
        ' �i�O���t��1���Ȃ��ꍇ�̃G���[������j
        On Error Resume Next
        ' ChartObjects�R���N�V����(�O���t)�S�̂��폜
        .ChartObjects.Delete
        ' �G���[�n���h�����O�����ɖ߂�
        On Error GoTo 0
        
        .Shapes.AddChart.Select
    End With
    With ActiveChart
        .SetSourceData _
            Source:=Range(Cells(1, 1), Cells(21, 2))
        .ChartType = xlColumnClustered
        .ChartTitle.Text = "�n�k�̔�����"
        With .Parent
            .Left = targetRange.Left
            .Top = targetRange.Top
            .Width = targetRange.Width
            .Height = targetRange.Height
        End With
    End With
    Cells(2, 1).Select
End Sub

