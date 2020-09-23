VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox Label1 
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "08:30:00"
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "OT Hours"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Working Hours"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Dim Str1$, Str2$, Str3$, Str4$
Text2.Text = DateTime.DateDiff("n", Text1.Text, Label1.Text)

Str1 = Text2.Text \ 60
Str2 = Text2.Text Mod 60
If Str1 > 24 Then
    Text3.Text = "More than a day.."
Else
    Text3.Text = Str1 & " hours & " & Str2 & " minutes"
    Text3.Text = Format(Text3.Text, "hh:mm")
    If Val(Text2.Text) > 480 Then
        Label2.Caption = Text2.Text - 480
        Str3 = Label2.Caption \ 60
        Str4 = Label2.Caption Mod 60
        Text4.Text = Str3 & " hours & " & Str4 & " minutes"
    End If
End If
End Sub

Private Sub Form_Load()
Text1.Text = Format(Text1.Text, "hh:mm:ss")
Label1.Text = Time
Label1.Text = Format(Label1.Text, "hh:mm:ss")

End Sub
