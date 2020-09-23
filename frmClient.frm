VERSION 5.00
Begin VB.Form frmClient 
   Caption         =   "Client Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   105
      TabIndex        =   2
      Top             =   2400
      Width           =   4500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   4485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "I'm a newly created Client Form :)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   4425
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        Me.Label1.Width = Me.Width - 330
        Me.Text1.Width = Me.Width - 330
        Me.Text1.Height = Me.Height - 1250
        Me.Text2.Width = Me.Text1.Width
        Me.Text2.Top = Me.Text1.Height + 450
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'this checks the array and finds the number associated
    'with this form
    Dim i As Long
    For i = LBound(sNewForm) To UBound(sNewForm)
        With sNewForm(i)
            If LCase(.Caption) = LCase(Me.Caption) Then
                'this is me in the array, set me
                'to nothing
                Debug.Print "Setting form array # " & i & " to nothing to be reused..."
                Set sNewForm(i) = Nothing
                Exit For
            End If
        End With
    Next i
End Sub
