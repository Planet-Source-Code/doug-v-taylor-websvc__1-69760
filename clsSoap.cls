VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'|--------------------------------------------------|
'| SOAP Authentication Header                       |
'|--------------------------------------------------|
Implements IHeaderHandler ' means this class is based on IHeaderHandler

' local variables to store properties
Private m_Username As String
Private m_Password As String

' the namespace which the service uses
Public NameSpace As String

Public Sub SetLogin(Username As String, Password As String)
    m_Username = Username
    m_Password = Password
End Sub

Private Sub iHeaderHandler_WriteHeaders(ByVal SoapSerializer As ISoapSerializer, ByVal obj As Object)
    'write the Authentication header
    With SoapSerializer
        ' start a parent Authentication element
        .StartHeaderElement "Authentication", NameSpace
            ' containing a username element with the username
            .StartElement "Username", NameSpace
                .WriteString m_Username
            .EndElement
            ' and a password element with the password
            .StartElement "Password", NameSpace
                .WriteString m_Password
            .EndElement
        .EndHeaderElement
    End With
End Sub

'some administration below

Private Function iHeaderHandler_ReadHeader(ByVal Reader As ISoapReader, ByVal HeaderNode As IXMLDOMNode, ByVal obj As Object) As Boolean
    iHeaderHandler_ReadHeader = False
End Function

Private Function iHeaderhandler_WillWriteHeaders() As Boolean
    iHeaderhandler_WillWriteHeaders = True
End Function

