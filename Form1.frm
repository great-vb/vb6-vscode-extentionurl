VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VSCode插件下载地址转换"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   10215
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1440
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下载"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   795
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1200
      Width           =   9975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   9975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1320
      TabIndex        =   4
      Top             =   75
      Width           =   105
   End
   Begin VB.Label Label2 
      Caption         =   "下载地址"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "扩展网址"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mWinHttpReq As New WinHttp.WinHttpRequest

Private Const URL_TEMPLETE = "https://${publisher}.gallery.vsassets.io/_apis/public/gallery/publisher/${publisher}/extension/${extension name}/${version}/assetbyname/Microsoft.VisualStudio.Services.VSIXPackage"
'https://marketplace.visualstudio.com/items?itemName=robertohuertasm.vscode-icons
Public Function DetailUrlToDownloadUrl(ByVal DetailUrl As String) As String
  Dim info() As String
  Dim downloadUrl As String
  Dim webCode As String
  Dim version As String
  
  info = Split(Split(DetailUrl, "=")(1), ".")
  webCode = GetHtml(DetailUrl)
  version = versionFromHtml(webCode)
  
  downloadUrl = Replace(URL_TEMPLETE, "${publisher}", info(0))
  downloadUrl = Replace(downloadUrl, "${extension name}", info(1))
  downloadUrl = Replace(downloadUrl, "${version}", version)
  
  Command1.Tag = info(1) & "-" & version
  
  DetailUrlToDownloadUrl = downloadUrl
End Function

Private Function BytesToBstr(strBody, CodeBase) '编码转换("UTF-8"或者"GB2312"或者"GBK")
  Dim ObjStream
  Set ObjStream = CreateObject("Adodb.Stream")
  With ObjStream
    .Type = 1
    .Mode = 3
    .Open
    .Write strBody
    .position = 0
    .Type = 2
    .Charset = CodeBase
    BytesToBstr = .ReadText
    .Close
  End With
  Set ObjStream = Nothing
End Function

Public Function GetHtml(ByVal URL As String) As String
  On Error GoTo ReStart
  Dim ExecTimes As Integer
  Dim htmlCode As String
  
  ExecTimes = 0
  
ReStart:
  If ExecTimes >= 3 Then
    GetHtml = ""
    Exit Function
  End If
  
  ExecTimes = ExecTimes + 1
  
  mWinHttpReq.Open "GET", URL, True
  mWinHttpReq.SetTimeouts 30000, 30000, 30000, 30000
  mWinHttpReq.setRequestHeader "Host", "marketplace.visualstudio.com"
  mWinHttpReq.setRequestHeader "Connection", "keep-alive"
  mWinHttpReq.setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
  mWinHttpReq.setRequestHeader "User-Agent", "Mozilla/5.0 (iPhone; U; CPU iPhone OS 4_2_1 like Mac OS X; en-us) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8C148 Safari/6533.18.5"
  mWinHttpReq.setRequestHeader "Accept-Language", "zh-CN,zh;q=0.8"
  mWinHttpReq.SetProxy 0
  mWinHttpReq.send            '发送
  mWinHttpReq.WaitForResponse '异步发送
  htmlCode = BytesToBstr(mWinHttpReq.responseBody, "UTF-8")
  GetHtml = htmlCode
End Function

Private Function versionFromHtml(ByVal HTML As String) As String
  Dim webCode As String
  webCode = Split(HTML, """version"":""")(1)
  webCode = Split(webCode, """")(0)
  versionFromHtml = webCode
End Function

Private Sub Command1_Click()
  On Error GoTo ExitPointX
  Label3.Caption = "尝试下载插件..."
  Dim YzmPic() As Byte

  mWinHttpReq.Open "GET", Text2.Text, True
  mWinHttpReq.Option(WinHttpRequestOption_EnableRedirects) = True
  mWinHttpReq.Option(WinHttpRequestOption_EnableHttpsToHttpRedirects) = True
  mWinHttpReq.SetTimeouts 30000, 30000, 30000, 30000
  mWinHttpReq.setRequestHeader "Host", Split(Split(Text2.Text, "//")(1), "/")(0)
  mWinHttpReq.setRequestHeader "Connection", "keep-alive"
  mWinHttpReq.setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
  mWinHttpReq.setRequestHeader "User-Agent", "Mozilla/5.0 (iPhone; U; CPU iPhone OS 4_2_1 like Mac OS X; en-us) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8C148 Safari/6533.18.5"
  mWinHttpReq.SetProxy 0
  mWinHttpReq.send            '发送
  mWinHttpReq.WaitForResponse '异步发送
  
  YzmPic = mWinHttpReq.responseBody
  Open App.Path & "\" & Command1.Tag & ".vsix" For Binary As #1
    Put #1, , YzmPic
  Close #1
  Label3.Caption = "下载完成"
  Timer1.Enabled = True
  Exit Sub
ExitPointX:
  Debug.Print Err.Description
  Label3.Caption = "下载出错"
  Timer1.Enabled = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Label3.Caption = "解析中..."
    DoEvents
    Text2.Text = DetailUrlToDownloadUrl(Text1.Text)
    Text1.Text = ""
    DoEvents
    Label3.Caption = ""
  End If
End Sub

Private Sub Timer1_Timer()
  Label3.Caption = ""
  Timer1.Enabled = False
End Sub
