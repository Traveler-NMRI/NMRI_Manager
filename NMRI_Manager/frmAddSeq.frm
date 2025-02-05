VERSION 5.00
Begin VB.Form frmAddSeq 
   Caption         =   "添加序列"
   ClientHeight    =   3780
   ClientLeft      =   7980
   ClientTop       =   3510
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   4620
   Begin VB.CommandButton Command2 
      Caption         =   "管理序列"
      Height          =   312
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   16
      Tag             =   "104"
      Top             =   3360
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
      Height          =   312
      Left            =   2160
      MaskColor       =   &H00000000&
      TabIndex        =   15
      Tag             =   "104"
      Top             =   3360
      Width           =   1092
   End
   Begin VB.CommandButton FinishCommand 
      Caption         =   "完成"
      Height          =   312
      Left            =   3360
      MaskColor       =   &H00000000&
      TabIndex        =   14
      Tag             =   "104"
      Top             =   3360
      Width           =   1092
   End
   Begin VB.Frame Frame2 
      Caption         =   "参数"
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4335
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   2880
         TabIndex        =   24
         Top             =   1410
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Left            =   1080
         TabIndex        =   23
         Top             =   1410
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   2880
         TabIndex        =   22
         Top             =   1050
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   1080
         TabIndex        =   21
         Top             =   1050
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   2880
         TabIndex        =   20
         Top             =   690
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   1080
         TabIndex        =   19
         Top             =   690
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   2880
         TabIndex        =   18
         Top             =   330
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1080
         TabIndex        =   17
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "#8 Time:"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "#7 Time:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "#6 Time:"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "#5 Time:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "#4 Time:"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "#3 Time:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "#2 Time:"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "#1 Time:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1320
         TabIndex        =   4
         Top             =   690
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   840
         TabIndex        =   2
         Top             =   320
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "序列集名称:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "序列:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAddSeq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isSpectroscopy As Boolean
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub FinishCommand_Click()
frmNew.ListView1.ListItems.Add(, , Str(frmNew.ListView1.ListItems.Count + 1)).SubItems(1) = Combo1.Text
Unload Me
End Sub
Public Function isSpecOrImage(ByVal isSpec As Boolean)
isSpectroscopy = isSpec
End Function

Private Sub Form_Load()
Combo1.Clear
If isSpectroscopy = True Then
For i = 1 To Val(Module1.ReadINI("GLOBAL", "COUNTS", "", ".\SEQUENCE.INI")) Step 1
If Module1.ReadINI("SEQUENCE" & Trim(Str(i)), "TYPE", "", ".\SEQUENCE.INI") = "SPECTROSCOPY" Then
Combo1.AddItem (Module1.ReadINI("SEQUENCE" & Trim(Str(i)), "NAME", "", ".\SEQUENCE.INI"))
End If
Next i
Else
For i = 1 To Val(Module1.ReadINI("GLOBAL", "COUNTS", "", ".\SEQUENCE.INI")) Step 1
If Module1.ReadINI("SEQUENCE" & Trim(Str(i)), "TYPE", "", ".\SEQUENCE.INI") = "IMAGING" Then
Combo1.AddItem (Module1.ReadINI("SEQUENCE" & Trim(Str(i)), "NAME", "", ".\SEQUENCE.INI"))
End If
Next i
End If
End Sub

