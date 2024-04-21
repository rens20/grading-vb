VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   8580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   13935
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   27
      ToolTipText     =   "4TH QRT"
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   26
      ToolTipText     =   "3RD QRT"
      Top             =   6240
      Width           =   1455
   End
   Begin VB.OptionButton Option8 
      BackColor       =   &H8000000D&
      Caption         =   "ICT PM 11-2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   5520
      TabIndex        =   25
      ToolTipText     =   "ICT PM 11-2"
      Top             =   2880
      Width           =   1695
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H8000000D&
      Caption         =   "ICT PM 11-6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5520
      TabIndex        =   24
      ToolTipText     =   "ICT PM 11-6"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H8000000D&
      Caption         =   "ICT AM 11-3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   7320
      TabIndex        =   23
      ToolTipText     =   "ICT AM 11-4"
      Top             =   2880
      Width           =   1695
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H8000000D&
      Caption         =   "ICT PM 11-4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   9120
      TabIndex        =   22
      ToolTipText     =   "ICT PM 11-4"
      Top             =   2880
      Width           =   1695
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H8000000D&
      Caption         =   "ICT PM 11-8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   9120
      TabIndex        =   21
      ToolTipText     =   "ICT PM 11-8"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H8000000D&
      Caption         =   "ICT AM 11-7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   7320
      TabIndex        =   20
      ToolTipText     =   "ICT AM 11-7"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H8000000D&
      Caption         =   "ICT AM 11-5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   3720
      TabIndex        =   19
      ToolTipText     =   "ICT AM 11-5"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000D&
      Caption         =   "ICT AM 11-1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3720
      TabIndex        =   18
      ToolTipText     =   "ICT AM 11-4"
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H8000000D&
      Caption         =   "about me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "about me"
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000D&
      Caption         =   "exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "exit"
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000D&
      Caption         =   "clear all"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "clear all"
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      TabIndex        =   13
      ToolTipText     =   "result"
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Enter  Your Grade ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   3600
      TabIndex        =   4
      Top             =   4200
      Width           =   5295
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000D&
         Caption         =   "clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "clear"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "average"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "average"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         ToolTipText     =   "2ND QRT"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         ToolTipText     =   "1ST QRT"
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "4rt Qtr:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "3rd Qtr:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "2nd Qtr:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "1st Qtr:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "[ Section ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   3600
      TabIndex        =   3
      Top             =   2640
      Width           =   7335
      Begin VB.CheckBox Check5 
         Caption         =   "Check1"
         Height          =   495
         Left            =   360
         TabIndex        =   17
         Top             =   2160
         Width           =   1695
      End
   End
   Begin VB.TextBox studentName 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      ToolTipText     =   "STUDENT NAME"
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Student's Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comprog Grading System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   3840
      TabIndex        =   0
      Top             =   480
      Width           =   8415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim grade1, grade2, grade3, grade4, average As Double

    ' Get grades from Textboxes
    grade1 = CDbl(Text2.Text)
    grade2 = CDbl(Text3.Text)
    grade3 = CDbl(Text4.Text)
    grade4 = CDbl(Text1.Text)

    ' Calculate average
    average = (grade1 + grade2 + grade3 + grade4) / 4

    ' Display average in Text6
    Text6.Text = CStr(average)
    
   
End Sub

Private Sub txtStudentName_Change()
    Dim studentName As String
    studentName = txtStudentName.Text
End Sub

Private Sub Command2_Click()

    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text1.Text = ""

    ' Clear result field
    Text6.Text = ""
End Sub

Private Sub Command3_Click()
studentName.Text = " "
   Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text1.Text = ""
    Text6.Text = ""
     Option1.Value = False
    
    Option2.Value = False
    Option3.Value = False
    Option4.Value = False
    Option5.Value = False
    Option6.Value = False
    Option7.Value = False
    Option8.Value = False
    
    




End Sub

Private Sub Command4_Click()
End

End Sub

Private Sub Command5_Click()
MsgBox " hello im rens acuna your humble fullstack developer"

End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Image1_Click()

End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    ' Check if the entered key is a number, a dot, or the Backspace key
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = 8 Then
        ' Allow the key press for numbers, dot, and Backspace
        If KeyAscii = 46 And InStr(Text2.Text, ".") > 0 Then ' Check Text2 for dot
            ' Cancel the key press if a dot already exists in the text
            KeyAscii = 0
        End If
    Else
        ' Cancel the key press for any other key
        KeyAscii = 0
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    ' Check if the entered key is a number, a dot, or the Backspace key
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = 8 Then
        ' Allow the key press for numbers, dot, and Backspace
        If KeyAscii = 46 And InStr(Text3.Text, ".") > 0 Then ' Check Text3 for dot
            ' Cancel the key press if a dot already exists in the text
            KeyAscii = 0
        End If
    Else
        ' Cancel the key press for any other key
        KeyAscii = 0
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    ' Check if the entered key is a number, a dot, or the Backspace key
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = 8 Then
        ' Allow the key press for numbers, dot, and Backspace
        If KeyAscii = 46 And InStr(Text4.Text, ".") > 0 Then ' Check Text4 for dot
            ' Cancel the key press if a dot already exists in the text
            KeyAscii = 0
        End If
    Else
        ' Cancel the key press for any other key
        KeyAscii = 0
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    ' Check if the entered key is a number, a dot, or the Backspace key
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = 8 Then
        ' Allow the key press for numbers, dot, and Backspace
        If KeyAscii = 46 And InStr(Text1.Text, ".") > 0 Then ' Check Text1 for dot
            ' Cancel the key press if a dot already exists in the text
            KeyAscii = 0
        End If
    Else
        ' Cancel the key press for any other key
        KeyAscii = 0
    End If
End Sub

