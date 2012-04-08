Attribute VB_Name = "WinInet"
Option Explicit
Option Private Module

'''@seealso http://msdn.microsoft.com/en-us/library/aa385483.aspx

''' InternetOpen AccessType
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0&
Private Const INTERNET_OPEN_TYPE_DIRECT = 1&
Private Const INTERNET_OPEN_TYPE_PROXY = 3&

''' InternetOpen Flags
Private Const INTERNET_FLAG_ASYNC = &H10000000
Private Const INTERNET_FLAG_FROM_CACHE = &H1000000
Private Const INTERNET_FLAG_OFFLINE = &H1000000

''' InternetOpenUrl Flags
Private Const INTERNET_FLAG_RELOAD = &H80000000

'''@return HINTERNET
'''@param  IN LPCTSTR Agent
'''@param  IN DWORD   AccessType
'''@param  IN LPCTSTR ProxyName
'''@param  IN LPCSTR  ProxyBypass
'''@param  IN DWORD   Flags
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" ( _
    ByVal lpAgent As String, _
    ByVal dwAccessType As Long, _
    ByVal lpProxyName As String, _
    ByVal lpProxyBypass As String, _
    ByVal dwFlags As Long _
    ) As Long

'''@return HINTERNET
'''@param  IN HINTERNET hInternet
'''@param  IN LPCTSTR   Url
'''@param  IN LPCTSTR   Headers OPTIONAL
'''@param  IN DWORD     HeadersSize
'''@param  IN DWORD     Flags
'''@param  IN DWORD     Context
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" ( _
    ByVal hInternet As Long, _
    ByVal lpUrl As String, _
    ByVal lpHeaders As String, _
    ByVal dwHeadersLength As Long, _
    ByVal dwFlags As Long, _
    ByVal dwContext As Long _
    ) As Long

'''@return BOOL
'''@param  IN  HINTERNET hFile
'''@param  IN  LPVOID    Buf
'''@param  IN  DWORD     BufSize
'''@param  OUT LPDWORD   ReadSize
Private Declare Function InternetReadFile Lib "wininet.dll" ( _
    ByVal hFile As Long, _
    ByRef lpBuffer As Byte, _
    ByVal dwBufSize As Long, _
    ByRef lpReadSize As Long _
    ) As Long

'''@return BOOL
'''@param  IN HINTERNET hInternet
Private Declare Function InternetCloseHandle Lib "wininet.dll" ( _
    ByVal hInternet As Long _
    ) As Long

Private Const AGENT_NAME = "Excel VBA"
Private Const DL_BUFFER_SIZE = 10240

Private Function FileNameInUrl(ByVal strUrl As String) As String
    Dim i As Integer
    
    i = InStrRev(strUrl, "/", Len(strUrl))
    If i > 0 Then strUrl = Mid(strUrl, i + 1)
    i = InStr(strUrl, "/")
    If i > 0 Then strUrl = Mid(strUrl, 1, i - 1)
    i = InStr(strUrl, "?")
    If i > 0 Then strUrl = Mid(strUrl, 1, i - 1)
    
    FileNameInUrl = strUrl
End Function

Public Function FileDownload(ByVal dlUrl As String, ByVal svPath As String) As String
    Dim hOpen As Long
    hOpen = InternetOpen(AGENT_NAME, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    FileDownload = FileDownloadImpl(hOpen, dlUrl, svPath)
    InternetCloseHandle hOpen
End Function
Private Function FileDownloadImpl( _
    ByVal hOpen As Long, ByVal dlUrl As String, ByVal svPath As String _
    ) As String
    
    If Right(svPath, 1) = "\" Then
        Dim fname As String: fname = FileNameInUrl(dlUrl)
        If Len(fname) > 0 Then
            svPath = Fso.BuildPath(svPath, fname)
        Else
            svPath = GetTempFilePath(svPath, ".html")
        End If
    End If
    
    Dim hConn As Long, succ As Boolean
    hConn = InternetOpenUrl(hOpen, dlUrl, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
    succ = hConn <> 0
    If Not succ Then GoTo Ending
    
    Dim strm As Object: Set strm = CreateAdoDbStream(adTypeBinary)
    strm.Open
    
    Dim sz As Long, diff As Long, buf(DL_BUFFER_SIZE - 1) As Byte
    Do
        InternetReadFile hConn, buf(0), DL_BUFFER_SIZE, sz
        If sz = 0 Then Exit Do
        
        strm.Write buf
        diff = DL_BUFFER_SIZE - sz
        If diff > 0 Then strm.Position = strm.Position - diff
    Loop
    strm.SetEOS
    
    strm.Position = 0
    strm.SaveToFile svPath, adSaveCreateOverWrite
    strm.Close
    
Ending:
    InternetCloseHandle hConn
    FileDownloadImpl = IIf(succ, svPath, vbNullString)
End Function
