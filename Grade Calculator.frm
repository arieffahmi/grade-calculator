VERSION 5.00
Begin VB.Form GradeCalculator 
   Caption         =   "Grade Calculator"
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   4020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculateGrade 
      Caption         =   "Calculate Grade"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtGrade 
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtScore 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblGrade 
      Caption         =   "Grade:"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblScore 
      Caption         =   "Score:"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "GradeCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCalculateGrade_Click()
    Dim intScore As Integer
    Dim strGrade As String
        
    intScore = Val(txtScore.Text)
    
    If IsNumeric(txtScore.Text) Then
        Select Case intScore
            Case 0 To 49
                strGrade = "F: Fail"
            Case 50 To 64
                strGrade = "C: Pass"
            Case 65 To 74
                strGrade = "B: Good"
            Case 75 To 89
                strGrade = "A: Very Good"
            Case 90 To 100
                strGrade = "A+: Excellent"
            Case Else
                MsgBox ("Invalid Input")
        End Select
    Else
        MsgBox ("Invalid Input")
    End If
    
    txtGrade.Text = strGrade
 
End Sub

