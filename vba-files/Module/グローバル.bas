Option Explicit

' �O���[�o���ϐ�
Public EXTRACT_SHEET As Worksheet
Public GRAPH_SHEET As Worksheet
' �O���[�o���萔
Public Const dateCol As Integer = 2 ' �N������
Public Const timeCol As Integer = 3 ' �����b��
Public Const locateCol As Integer = 23 ' �k���n����

' �I�u�W�F�N�g���
Public Sub setObj()
    Set EXTRACT_SHEET = ThisWorkbook.Worksheets("���o")
    Set GRAPH_SHEET = ThisWorkbook.Worksheets("�O���t")
End Sub

' �I�u�W�F�N�g�J��
Public Sub releaseObj()
    Set EXTRACT_SHEET = Nothing
    Set GRAPH_SHEET = Nothing
End Sub

