VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CdoMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''+----                                                                   --+
'''|                             Ariawase 0.6.0 Beta                         |
'''|                Ariawase is free library for VBA cowboys.                |
'''|           The Project Page: https://github.com/igeta/Ariawase           |
'''+--                                                                   ----+
Option Explicit

''' @seealso http://msdn.microsoft.com/en-us/library/ms526266.aspx

Private Enum CdoSendUsing
    cdoSendUsingPickup = 1
    cdoSendUsingPort = 2
End Enum

''' @seealso http://msdn.microsoft.com/en-us/library/ms872853.aspx
Private Const uriCdoConf As String = "http://schemas.microsoft.com/cdo/configuration/"
Private Const cdoSendUsingMethod  As String = uriCdoConf & "sendusing"
Private Const cdoSMTPServer       As String = uriCdoConf & "smtpserver"
Private Const cdoSMTPServerPort   As String = uriCdoConf & "smtpserverport"
Private Const cdoSMTPAuthenticate As String = uriCdoConf & "smtpauthenticate"
Private Const cdoSMTPUseSSL       As String = uriCdoConf & "smtpusessl"
Private Const cdoSendUserName     As String = uriCdoConf & "sendusername"
Private Const cdoSendPassword     As String = uriCdoConf & "sendpassword"
Private Const cdoLanguageCode     As String = uriCdoConf & "languagecode"

Private xMyMailAddr As String
Private xConf As Object 'Is CDO.Configuration

Private xMsg As Object 'Is CDO.Message

Public Sub Configure( _
    ByVal myMailAddr As String, _
    Optional ByVal smtpServer As String, _
    Optional ByVal smtpAuth As CdoProtocolsAuthentication = cdoAnonymous, _
    Optional ByVal sndUser As String, _
    Optional ByVal sndPassword As String, _
    Optional ByVal useSSL As Boolean = False _
    )
    
    xMyMailAddr = myMailAddr
    
    Dim smtpAddrPort As Variant: smtpAddrPort = Split(smtpServer, ":")
    If UBound(smtpAddrPort) < 0 Then
        Set xConf = Nothing
        GoTo Escape
    End If
    
    Set xConf = CreateObject("CDO.Configuration")
    With xConf.Fields
        .Item(cdoSendUsingMethod) = CdoSendUsing.cdoSendUsingPort
        .Item(cdoSMTPServer) = smtpAddrPort(0)
        Select Case UBound(smtpAddrPort)
            Case 0: .Item(cdoSMTPServerPort) = 25
            Case 1: .Item(cdoSMTPServerPort) = smtpAddrPort(1)
            Case Else: Err.Raise 5
        End Select
        
        .Item(cdoSMTPAuthenticate) = smtpAuth
        .Item(cdoSendUserName) = sndUser
        .Item(cdoSendPassword) = sndPassword
        .Item(cdoSMTPUseSSL) = useSSL
    End With
    xConf.Fields.Update
    
Escape:
End Sub

Public Sub NewMail()
    If Len(xMyMailAddr) < 1 Then Err.Raise 5
    
    Set xMsg = CreateObject("CDO.Message")
    Set xMsg.Configuration = xConf
    'xMsg.Fields.Item("urn:schemas:mailheader:X-Mailer") = "Microsoft CDO for Windows 2000"
    'xMsg.Fields.Update()
    xMsg.From = xMyMailAddr
End Sub

Private Function GetAddrList(ByVal addlst As Variant) As String
    If IsArray(addlst) Then GetAddrList = Join(addlst, "; ") Else GetAddrList = CStr(addlst)
End Function

Public Sub SetAddress( _
    ByVal addrTo As Variant, _
    Optional ByVal addrCC As Variant = Empty, _
    Optional ByVal addrBCC As Variant = Empty _
    )
    
    xMsg.To = GetAddrList(addrTo)
    xMsg.CC = GetAddrList(addrCC)
    xMsg.BCC = GetAddrList(addrBCC)
End Sub

Public Sub SetMessage( _
    ByVal msgSubject As String, ByVal msgBody As String, _
    Optional ByVal chrset As String = cdoISO_2022_JP, _
    Optional ByVal contTransEnc As String = cdoBase64 _
    )
    
    xMsg.Subject = msgSubject
    xMsg.TextBody = msgBody
    xMsg.TextBodyPart.Charset = chrset
    xMsg.TextBodyPart.ContentTransferEncoding = contTransEnc
End Sub

Public Sub AddAttachments(ByVal atchFiles As Variant)
    Dim fpath As Variant '(Of String)
    For Each fpath In atchFiles: xMsg.AddAttachment fpath: Next
End Sub

Public Sub SendMail()
    xMsg.Send
End Sub

Public Sub QuickSendMail( _
    ByVal addrTo As Variant, ByVal msgSubject As String, ByVal msgBody As String _
    )
    
    If Len(xMyMailAddr) < 1 Then Err.Raise 5
    
    Dim msg As Object: Set msg = CreateObject("CDO.Message")
    Set msg.Configuration = xConf
    msg.From = xMyMailAddr
    
    msg.To = GetAddrList(addrTo)
    msg.Subject = msgSubject
    msg.TextBody = msgBody
    
    msg.Send
End Sub
