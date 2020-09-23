VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "File Transmission to Web Service"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txt 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   12
      Text            =   "Txt(4)"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   3
      Left            =   600
      TabIndex        =   10
      Text            =   "Txt(3)"
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Txt 
      Height          =   855
      Index           =   5
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "fMain.frx":0000
      Top             =   2760
      Width           =   5895
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   4
      Text            =   "Txt(2)"
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Text            =   "Txt(1)"
      Top             =   0
      Width           =   3735
   End
   Begin VB.TextBox Txt 
      Height          =   855
      Index           =   0
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "fMain.frx":0009
      Top             =   1920
      Width           =   5895
   End
   Begin VB.CommandButton cStop 
      Caption         =   "Stop"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton cStart 
      Caption         =   "Start"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Lbl 
      Caption         =   "Password:"
      Height          =   375
      Index           =   4
      Left            =   3120
      TabIndex        =   11
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Lbl 
      Caption         =   "UserID:"
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Lbl 
      Caption         =   "Process Files Like"
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Lbl 
      Caption         =   "Web Service Address"
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Lbl 
      Caption         =   "File:"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   4575
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WebAdr$
Private FName$
Private bUpdate As Boolean
Private CurFName$
Private IP$
Private dStart As Date
Private dFStart As Date
Private lBytes As Long
Private sBytes As Single
Private sFiles As Long, sRecs As Long
Private Stopit As Boolean
Private XFRErr As Boolean
Private SoapClnt As New SoapClient30
Private UserID$, PW$

Private Sub StatusUpdate()
Dim x$
x$ = "Started: " + Format$(dStart) + vbCrLf
x$ = x$ + "Files Transferred: " + Format$(sFiles)
Txt(5) = x$
Txt(5).Refresh
End Sub
Private Function MkName$(Mask$)
Dim i As Integer
Dim y$
y$ = Date$ + Time$
i = InStr(Mask$, ".")
If i > 0 Then
 MkName$ = Left$(Mask$, i - 1) + y$ + Mid$(Mask$, i)
Else
 MkName$ = Mask$ + y$
End If
End Function

Private Sub cStart_Click(Index As Integer)
Dim SoapCls As New clsSoap

On Error GoTo estart
sFiles = 0
sRecs = 0
Set SoapClnt = New SoapClient30
Close #1
Close #2
Open App.Path + "\WebSvc.LOG" For Append As #2
If Len(WebAdr$) = 0 Then Txt(1).SetFocus
If Len(FName$) = 0 Then Txt(2).SetFocus
If Len(UserID$) = 0 Then Txt(3).SetFocus
If Len(PW$) = 0 Then Txt(4).SetFocus
'If fMain.ActiveControl.Caption <> "Start" Then Exit Sub
Txt(0) = ""
If Len(WebAdr$) = 0 Or Len(FName$) = 0 Or Len(UserID$) = 0 Or Len(PW$) = 0 Then
 Txt(0) = "Missing Some Data field needed to Start" + vbCrLf + "Enter Missing Data and Press Start Again"
 Exit Sub
End If

Txt(0) = ""
Print #2, ""
Print #2, ""
Print #2, "Started: " + Format$(Now) + ", attempting Web Service Connect to " + WebAdr$
Print #2, "Using Logon: " + UserID$ + " and Password: " + PW$
SoapInit
If InStr(Txt(0), " Not ") Then
 Print #2, Txt(0)
 Print #2, ""
 Print #2, "Stopped for connection error at: " + Format$(Now)
Exit Sub
End If
Process
Close #1
'Close #2
dStart = Now
dFStart = Now
lBytes = 0
estart:
End Sub
Private Sub Process()
Dim Path$, UseFile$, MaskFile$, y$, Fil$, Rslt$

  
  MaskFile$ = Txt(2).Text
  UseFile$ = MaskFile$
  Path$ = MaskFile$
  While InStr("\:", Right$(Path$, 1)) = 0 And Len(Path$) > 0
   Path$ = Left$(Path$, Len(Path$) - 1)
  Wend
  If Len(Dir$(MaskFile$)) = 0 And (InStr(MaskFile$, "*") + InStr(MaskFile$, "?") = 0) Then
      GoTo Xit
  Else
    UseFile$ = Dir$(MaskFile$)
    Txt(0).Text = "Waiting for " + MaskFile$
  End If
  While Not Stopit
    While Len(UseFile$) = 0
     Txt(0).Text = "Waiting for " + MaskFile$
     DoEvents
     If Stopit Then GoTo Xit  'this assures a log file is created
     UseFile$ = Dir$(MaskFile$)
     DoEvents
    Wend
    locfile$ = Path$ + UseFile$
    XFRErr = False
    FTPFile$ = UseFile$
    Txt(0).Text = "Sending " + locfile$
    Print #2, Txt(0).Text
    Close #1
    Open locfile$ For Input As #1
    sRecs = 0
    Do While Not EOF(1)
     Line Input #1, Fil$
     Rslt$ = SendData(Fil$)
     DoEvents
     If XFRErr Then Exit Do
     sRecs = sRecs + 1
    Loop
    Close #1
    If Not xrferr Then sFiles = sFiles + 1
    StatusUpdate
    fMain.Refresh
    If Not XFRErr Then
      y$ = UseFile$        'make new file name in y$
      While Right$(y$, 1) <> "." And Len(y$) > 0
       y$ = Left$(y$, Len(y$) - 1)
      Wend
      y$ = "Done-" + Format$(Now, "Medium Date") + "-" + Time$ + "-" + y$ + "DUN"
      y$ = SubTran$(y$, ":", "-")
      y$ = SubTran$(y$, "/", "-")
      If Len(Dir$(Path$ + y$)) > 0 Then Kill Path$ + y$
      Name Path$ + UseFile$ As Path$ + y$ 'prevent from being read again
      Print #2, Format$(Now) + ": Processed " + Format$(sRecs) + " From " + Path$ + UseFile$
    Else
      Print #2, Format$(Now) + ": Error Processing " + Path$ + UseFile$
      GoTo Xit
    End If
    UseFile$ = ""
    fMain.Refresh
  Wend
'Close #2
Xit:
 cStop_Click (0)
 End
End Sub

Private Sub cStop_Click(Index As Integer)
 Dim x$
 Close #2
 Open App.Path + "\WebSvc.LOG" For Append As #2
 Stopit = True
 If Index = 1 Then
  Print #2, "Stopped by Stop Button at: " + Format$(Now)
 Else
  Print #2, "Stopped by Program Logic at: " + Format$(Now)
 End If
 If bUpdate Then
  x$ = App.Path + "\" + "WebSvc.Set"
  If Len(Dir$(x$)) Then Kill x$
  Close #1
  Open x$ For Output As #1
  Print #1, WebAdr$
  Print #1, UserID$
  Print #1, PW$
  Print #1, FName$
  Close #1
 End If
 End

End Sub

Private Sub Form_Load()
Dim x$
 On Error Resume Next
 dStart = Now
 sBytes = 0
 sFiles = 0
 bUpdate = False
 Txt(0) = ""
 x$ = App.Path + "\" + "WebSvc.Set"
 If Len(Dir$(x$)) Then
  Open x$ For Input As #1
  Line Input #1, WebAdr$
  Txt(1) = WebAdr$
  Line Input #1, UserID$
  Txt(3) = UserID$
  Line Input #1, PW$
  Txt(4) = PW$
  Line Input #1, FName$
  Txt(2) = FName$
  Close #1
 End If
 StatusUpdate
 fMain.Refresh
End Sub

Private Sub Form_Resize()
Txt(1).Move Txt(1).Left, Txt(1).Top, Me.Width - Txt(1).Left - 100
Txt(2).Move Txt(2).Left, Txt(2).Top, Me.Width - Txt(2).Left - 100
cStart(0).Move cStart(0).Left, Me.Height - cStart(0).Height - 650
cStop(1).Move cStop(1).Left, cStart(0).Top
Txt(0).Move Txt(0).Left, Txt(0).Top, Me.Width - Txt(0).Left - 200, (Me.Height - Txt(0).Top - 800 - cStart(0).Height) / 2
Txt(5).Move Txt(5).Left, Txt(0).Top + Txt(0).Height, Txt(0).Width, Txt(0).Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
 cStop_Click (Cancel)
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Select Case Index
 Case 1
  If WebAdr$ <> Txt(1) Then
   WebAdr$ = Txt(1)
   bUpdate = True
  End If
 Case 2
  If FName$ <> Txt(2) Then
   FName$ = Txt(2)
   bUpdate = True
  End If
 Case 3
  If UserID$ <> Txt(3) Then
   UserID$ = Txt(3)
   bUpdate = True
  End If
 Case 4
  If PW$ <> Txt(4) Then
   PW$ = Txt(4)
   bUpdate = True
  End If
End Select
End Sub
Public Function SubTran$(Inst$, Lookfor$, ByVal Change2$)
 Dim i As Long, j As Long
 Dim InString$
 InString$ = Inst$
 j = 1
 i = InStr(j, InString$, Lookfor$)
 While i > 0
   InString$ = Left$(InString$, i - 1) + Change2$ + Mid$(InString$, i + Len(Lookfor$))
   j = i + Len(Change2$)
   i = InStr(j, InString$, Lookfor$)
 Wend
 SubTran$ = InString$
End Function
Private Function SendData$(ByVal MedData As String)


    On Error Resume Next
    
    XFRErr = False

    SendData$ = UCase$(SoapClnt.PatientUpdate2(MedData))
    If Err.Number <> 0 Then
     Txt(0) = "Error Transmitting Record!" + vbCrLf + Err.Description + vbCrLf + MedData
     Print #2, Txt(0)
     XFRErr = True
    ElseIf SendData$ = "FAIL" Then
     Txt(0) = "Record Transmitted but Not Acceptable to ProCareRX!" + vbCrLf + MedData
     Print #2, Txt(0)
     XFRErr = True
    ElseIf SendData$ = "PASS" Then
     Txt(0) = "Record " + Format$(sRecs) + " Accepted by ProCareRX!"
     Txt(0).Refresh
    End If
    
End Function

Private Sub SoapInit()
Dim Auth$, Base64Auth$
Dim SoapCls As New clsSoap

On Error GoTo SCErr

 Txt(0) = ""
 Auth$ = LCase$(WebAdr$ + "?wsdl")
 Auth$ = SubTran$(Auth$, "https://", "https://" + UserID$ + ":" + PW$ + "@")
    Print #2, "Using Connect URL: " + Auth$

    SoapClnt.MSSoapInit Auth$
    
    SoapCls.NameSpace = "PutPBM"
    SoapCls.SetLogin UserID$, PW$
    Set SoapClnt.HeaderHandler = SoapCls

    Print #2, "Connection SUccessful at " + Format$(Now)

Exit Sub
SCErr:
    
    If Err.Number <> 0 Then Txt(0) = "Soap Connection Not Established!" + vbCrLf + Err.Description
    'SOAPClnt.ConnectorProperty("AuthUser") = UserID$
    'SOAPClnt.ConnectorProperty("AuthPassword") = PW$
     
End Sub


