VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmTest 
   Caption         =   "Testing Chip8086XT"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   5685
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load file"
      Height          =   315
      Left            =   5475
      TabIndex        =   27
      Top             =   1695
      Width           =   1230
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save File"
      Height          =   315
      Left            =   4155
      TabIndex        =   26
      Top             =   1695
      Width           =   1230
   End
   Begin RichTextLib.RichTextBox RtxtResult 
      Height          =   1860
      Left            =   75
      TabIndex        =   10
      Top             =   2070
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3281
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmTest.frx":0CCA
   End
   Begin RichTextLib.RichTextBox RtxtSNum 
      Height          =   510
      Left            =   1395
      TabIndex        =   4
      Top             =   630
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   900
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmTest.frx":0D4C
   End
   Begin RichTextLib.RichTextBox RtxtFNum 
      Height          =   510
      Left            =   1395
      TabIndex        =   2
      Top             =   90
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   900
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmTest.frx":0DCC
   End
   Begin VB.CommandButton cmdClean 
      Caption         =   "Clea&n"
      Height          =   315
      Left            =   4155
      TabIndex        =   9
      Top             =   1275
      Width           =   1230
   End
   Begin VB.Frame fraTime 
      Caption         =   "Timing in Opration (in sec)"
      Height          =   1590
      Left            =   105
      TabIndex        =   12
      Top             =   3975
      Width           =   6975
      Begin VB.Label lblTime 
         Caption         =   "0"
         Height          =   315
         Index           =   5
         Left            =   4515
         TabIndex        =   24
         Top             =   1020
         Width           =   1245
      End
      Begin VB.Label lblTime 
         Caption         =   "0"
         Height          =   315
         Index           =   4
         Left            =   3075
         TabIndex        =   23
         Top             =   1020
         Width           =   1245
      End
      Begin VB.Label lblTime 
         Caption         =   "0"
         Height          =   315
         Index           =   3
         Left            =   1605
         TabIndex        =   22
         Top             =   1020
         Width           =   1245
      End
      Begin VB.Label lblTime 
         Caption         =   "0"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   1020
         Width           =   1245
      End
      Begin VB.Label Label6 
         Caption         =   "Division"
         Height          =   270
         Left            =   4500
         TabIndex        =   20
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label5 
         Caption         =   "Multiplication"
         Height          =   255
         Left            =   3060
         TabIndex        =   19
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label4 
         Caption         =   "Subtration"
         Height          =   210
         Left            =   1605
         TabIndex        =   18
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "Addtion"
         Height          =   255
         Left            =   165
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblTime 
         Caption         =   "0"
         Height          =   240
         Index           =   1
         Left            =   5205
         TabIndex        =   16
         Top             =   285
         Width           =   1725
      End
      Begin VB.Label lblTime 
         Caption         =   "0"
         Height          =   225
         Index           =   0
         Left            =   1800
         TabIndex        =   15
         Top             =   285
         Width           =   1470
      End
      Begin VB.Label Label2 
         Caption         =   "Second Number Initialize:"
         Height          =   225
         Left            =   3390
         TabIndex        =   14
         Top             =   270
         Width           =   1845
      End
      Begin VB.Label Label1 
         Caption         =   "First Number Initialize:"
         Height          =   210
         Left            =   165
         TabIndex        =   13
         Top             =   270
         Width           =   1590
      End
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   2610
      TabIndex        =   8
      Top             =   1275
      Width           =   465
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   1935
      TabIndex        =   7
      Top             =   1275
      Width           =   465
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   1260
      TabIndex        =   6
      Top             =   1275
      Width           =   465
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   615
      TabIndex        =   5
      Top             =   1275
      Width           =   465
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   5475
      TabIndex        =   0
      Top             =   1275
      Width           =   1230
   End
   Begin VB.Label lblResult 
      Height          =   210
      Left            =   150
      TabIndex        =   25
      Top             =   1755
      Width           =   3855
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6765
      TabIndex        =   11
      ToolTipText     =   "About"
      Top             =   1530
      Width           =   375
   End
   Begin VB.Label lblSNum 
      Alignment       =   1  'Right Justify
      Caption         =   "Second Number"
      Height          =   240
      Left            =   90
      TabIndex        =   3
      ToolTipText     =   "Click here to generate Random Numbers"
      Top             =   765
      Width           =   1245
   End
   Begin VB.Label lblFnum 
      Alignment       =   1  'Right Justify
      Caption         =   "First Number"
      Height          =   255
      Left            =   195
      TabIndex        =   1
      ToolTipText     =   "Click here to generate Random Numbers"
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim myObj As MathOperator.Operations
Dim myObj As Chip8086XT.MathOperator
Private Const Fnum As String = "\FirstNum.rtf"
Private Const SNum As String = "\SecondNum.rtf"
Private Const RNum As String = "\Result.rtf"


Private Sub cmdClean_Click()
Dim i As Byte
RtxtFNum.Text = "0"
RtxtSNum.Text = "0"
RtxtResult.Text = "0"
 For i = 0 To lblTime.UBound
  lblTime(i).Caption = "0"
 Next i
End Sub

Private Sub cmdClose_Click()
    Set myObj = Nothing
    End
End Sub

Private Sub cmdLoad_Click()
On Error GoTo ErrMsg
    RtxtFNum.LoadFile App.Path & Fnum ', rtfRTF
    RtxtSNum.LoadFile App.Path & SNum ', rtfRTF
    RtxtResult.LoadFile App.Path & RNum ', rtfRTF
Exit Sub
ErrMsg:
    MsgBox "The error in opening the files. Error Number is = " & Err.Number & " Error Description = " & Err.Description
End Sub

Private Sub cmdSave_Click()
    RtxtFNum.SaveFile App.Path & Fnum ', rtfRTF
    RtxtSNum.SaveFile App.Path & SNum ', rtfRTF
    RtxtResult.SaveFile App.Path & RNum ', rtfRTF
End Sub

Private Sub Form_Load()
    Set myObj = New Chip8086XT.MathOperator
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblFnum.Font.Underline = False
    lblSNum.Font.Underline = False
    lblAbout.Font.Underline = False
End Sub

Private Sub lblAbout_Click()
    MsgBox myObj.About, vbInformation, "Thanks for the Testing"
   ' txtFNum.Text = myObj.About
End Sub

Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblAbout.Font.Underline = True
End Sub

Private Sub lblFnum_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblFnum.Font.Underline = True
End Sub
Private Sub RtxtFNum_GotFocus()
    RtxtFNum.SelStart = 0
    RtxtFNum.SelLength = Len(RtxtFNum.Text)
End Sub

Private Sub RtxtFNum_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 RtxtFNum.ToolTipText = "Total No of Digits: " & CStr(Len(RtxtFNum.Text))
End Sub

Private Sub RtxtSNum_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 RtxtSNum.ToolTipText = "Total No. of Digits: " & CStr(Len(RtxtSNum.Text))
End Sub

Private Sub RtxtResult_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 RtxtResult.ToolTipText = "Total No. of Digits: " & CStr(Len(RtxtResult.Text))
End Sub

Private Sub rtxtSNum_GotFocus()
    RtxtSNum.SelStart = 0
    RtxtSNum.SelLength = Len(RtxtSNum.Text)
End Sub

Private Sub lblSNum_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblSNum.Font.Underline = True
End Sub

Private Sub lblFnum_Click()
Dim i As Variant
i = InputBox("How many digit you want to creat?", "Create First Random Numbers")
    If i <> "" Then CreatRndNum 1, Filter(i)
End Sub

Private Sub lblSNum_Click()
Dim i As Variant
i = InputBox("How many digit you want to creat?", "Create Second Random Numbers")
    If i > 0 Then CreatRndNum 2, Filter(i)
End Sub

Function Filter(ByVal vNumber As String) As String
On Error GoTo ErrMsg
 Dim vReturn As String, i As Variant
 Dim SearchingString As String
 SearchingString = "0123456789."
  vReturn = ""
 For i = 0 To Len(vNumber) - 1
  If InStr(SearchingString, Mid(vNumber, i + 1, 1)) > 0 Then
       vReturn = vReturn + Mid(vNumber, i + 1, 1)
  End If
 Next i
 If vReturn = "" Then vReturn = "0"
  Filter = vReturn
 Exit Function

ErrMsg:
 Filter = "0"
 MsgBox "The Error No is = " & CStr(Err.Number) & " Error Description = " & Err.Description
End Function

Public Function CreatRndNum(ByVal vWhere As Byte, ByVal vHowMany As Double)
Dim i As Variant, vStr As String
vStr = ""

 For i = 1 To vHowMany
    vStr = vStr + CStr(Int(Rnd(1) * 10))
 Next i
'MsgBox i
    Select Case vWhere
 Case 1
 If vStr = "" Then
    RtxtFNum = "0"
 Else
    RtxtFNum.Text = vStr
 End If
 Case 2
    RtxtSNum.Text = vStr
    End Select
    
End Function
Private Sub cmdOpt_Click(Index As Integer)
    
    myObj.FirstNum = RtxtFNum.Text
    lblTime(0).Caption = myObj.ExpeledTime 'In sec
    myObj.SecondNum = RtxtSNum.Text
    lblTime(1).Caption = myObj.ExpeledTime  'FormatDateTime(myObj.ExpeledTime)

Select Case Index
    Case 0
        myObj.Addtion
        lblResult.Caption = "Result After Addition"
        lblTime(2).Caption = myObj.ExpeledTime  'FormatDateTime(myObj.ExpeledTime)
    Case 1
        myObj.Subtraction
        lblResult.Caption = "Result After Subtraction"
        lblTime(3).Caption = myObj.ExpeledTime  'FormatDateTime(myObj.ExpeledTime)
    Case 2
        myObj.Multiply
        lblResult.Caption = "Result After Multiplication"
        lblTime(4).Caption = myObj.ExpeledTime
    Case 3
        myObj.Division
        lblResult.Caption = "Result After division"
        lblTime(5).Caption = myObj.ExpeledTime
End Select

    If myObj.ErrNum = 0 Then
        RtxtResult.Text = myObj.Result
    Else
        MsgBox myObj.ErrDescription
    End If
End Sub

