Attribute VB_Name = "Resource"
''' +----                                                             --+ '''
''' |                          Ariawase 0.5.0                           | '''
''' |             Ariawase is free library for VBA cowboys.             | '''
''' |        The Project Page: https://github.com/igeta/Ariawase        | '''
''' +--                                                             ----+ '''
Option Explicit
Option Private Module

Public Enum CdoProtocolsAuthentication
    cdoAnonymous = 0
    cdoBasic = 1
    cdoNTLM = 2
End Enum

Public Const cdo7bit        As String = "7bit"
Public Const cdo8bit        As String = "8bit"
Public Const cdoISO_2022_JP As String = "iso-2022-jp"
Public Const cdoShift_JIS   As String = "shift-jis"
Public Const cdoEUC_JP      As String = "euc-jp"
Public Const cdoUTF_8       As String = "utf-8"

Public Const cdoBase64          As String = "base64"
Public Const cdoQuotedPrintable As String = "quoted-printable"
