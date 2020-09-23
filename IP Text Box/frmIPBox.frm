VERSION 5.00
Begin VB.Form frmIPBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advance TCP/IP Text Box"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3705
   Icon            =   "frmIPBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   1350
      TabIndex        =   8
      Top             =   1050
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter IP Address"
      Height          =   915
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3585
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   510
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "0"
         ToolTipText     =   "Enter Your Server IP Address Located at the Lower Right of PC-Time Logger"
         Top             =   420
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Enter Your Server IP Address Located at the Lower Right of PC-Time Logger"
         Top             =   420
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1890
         MaxLength       =   3
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Enter Your Server IP Address Located at the Lower Right of PC-Time Logger"
         Top             =   420
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   2580
         MaxLength       =   3
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Enter Your Server IP Address Located at the Lower Right of PC-Time Logger"
         Top             =   420
         Width           =   555
      End
      Begin VB.TextBox Sepa1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   1050
         TabIndex        =   3
         Text            =   "."
         Top             =   420
         Width           =   165
      End
      Begin VB.TextBox Sepa1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   2430
         TabIndex        =   2
         Text            =   "."
         Top             =   420
         Width           =   165
      End
      Begin VB.TextBox Sepa1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   1740
         TabIndex        =   1
         Text            =   "."
         Top             =   420
         Width           =   165
      End
      Begin VB.Shape Shape1 
         Height          =   345
         Left            =   480
         Top             =   390
         Width           =   2685
      End
   End
End
Attribute VB_Name = "frmIPBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Text1_Change(Index As Integer)
If Index = 1 Then
    If Val(Text1(Index)) > 223 Then
        MsgBox Text1(Index) & " is not a valid entry. Please specify a value between 1 to 223", vbOKOnly + vbExclamation, "Error"
        Text1(Index) = 223
        Text1(Index).SetFocus
    End If
Else
    If Val(Text1(Index)) > 255 Then
        MsgBox Text1(Index) & " is not a valid entry. Please specify a value between 1 to 255", vbOKOnly + vbExclamation, "Error"
        Text1(Index) = 255
        Text1(Index).SetFocus
    End If
End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = 8 Then
    If Text1(Index).SelStart = 0 Then
        If Index > 1 Then
            Text1(Index - 1).SetFocus
            Text1(Index - 1).SelStart = Len(Text1(Index - 1))
            Text1(Index - 1).SelLength = 0
        End If
    End If
End If
If KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    If Text1(Index).SelStart = Len(Text1(Index)) Then
        If Index < 4 Then
            Text1(Index + 1).SetFocus
            Text1(Index + 1).SelStart = 0
            Text1(Index + 1).SelLength = 0
        End If
    End If
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 0
    If Index < 4 Then
        Text1(Index + 1).SetFocus
        Text1(Index + 1).SelStart = 0
        Text1(Index + 1).SelLength = Len(Text1(Index + 1))
    End If
End If
If KeyAscii = 8 Then
    Exit Sub
End If
If KeyAscii < 48 Or KeyAscii > 57 Then
       KeyAscii = 0
End If

If Index < 4 And Len(Text1(Index)) >= 2 And Text1(Index).SelText <> Text1(Index) Then
    If Text1(Index).SelStart = Len(Text1(Index)) Then
        Text1(Index + 1).SetFocus
        Text1(Index + 1).SelStart = 0
        Text1(Index + 1).SelLength = Len(Text1(Index + 1))
    End If
End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index) = Val(Text1(Index))
End Sub
