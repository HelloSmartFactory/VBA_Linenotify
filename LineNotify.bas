Attribute VB_Name = "LineNotify"
Option Explicit

'token�̐ݒ�
Const strToken As String = "yourtoken"


Sub doMsg()
    Dim msg As String
    msg = "2020/06/12"
    sendLineNotify (msg)
End Sub


'Line���M
'�����Fstr���b�Z�[�W
Private Sub sendLineNotify(msg As String)
    '�I�u�W�F�N�g���� '�Q�Ɛݒ�Ȃ�
    Dim objHTTP As Object
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    '�Q�Ɛݒ肠��̏ꍇ�͂����� microsoft xml v6.0
    'Dim objHTTP As XMLHTTP60
    'Set objHTTP = New XMLHTTP60
    
On Error GoTo errHandler    '�G���[�͔�΂�
    objHTTP.Open "POST", "https://notify-api.line.me/api/notify", False '�I�u�W�F�N�g������
    '�w�b�_�ݒ�
    objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" '��͕��@�w��
    objHTTP.setRequestHeader "Authorization", "Bearer " + strToken  'headers
    '���M
    objHTTP.send "message=" + msg '+ "&stickerPackageId=1" + "&stickerId=113" 'payload

    '�X�e�[�^�X�m�F
    If objHTTP.Status = 200 Then    '400���N�G�X�g���s�� 401�A�N�Z�X�g�[�N���������@500�T�[�o���G���[�ɂ�莸�s
        Debug.Print "���܂��������� " + objHTTP.responseText
    Else
        Debug.Print "�Ȃ񂩂��������ŁI�@" + objHTTP.responseText
    End If
    
    Set objHTTP = Nothing   '�I�u�W�F�N�g�j��
Exit Sub

errHandler: '�G���[�Ŕ��ł���
    Dim number As Long: number = Err.number '�G���[�R�[�h�擾
    
    '�G���[�ʂɏ���
    Select Case number
        Case -2146197211
            Debug.Print "�G���[�R�[�h = " & number & vbCrLf & "     �w�肳�ꂽ�\�[�X��������܂���B�l�b�g���[�N�ڑ����m�F���Ă��������B"
        Case Else
            Debug.Print "�G���[�R�[�h = " & number
    End Select
    
    Set objHTTP = Nothing   '�I�u�W�F�N�g�j��
End Sub
