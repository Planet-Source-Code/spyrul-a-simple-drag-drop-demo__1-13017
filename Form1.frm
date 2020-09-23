VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "Drag Drop Demo"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Drag Me!"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Drag over these three colors and the control will lock into position."
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Drag Drop Demo
' Drag the command button all over the form and it will stay where you leave it
' Drag the button over the different labels to lock them into that position
' Simple demo to show how to get things to drag and how to get them to drop
'   and lock.

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Drag vbBeginDrag ' Starts drag
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Drag vbEndDrag ' Ends Drag
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
Source.Move (X - Source.Width / 2), (Y - Source.Height / 2) ' Leaves the object
' where you want on the form.
End Sub

Private Sub Label1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Source.Left = Label1.Left  'Locks the control into position over Label1
Source.Top = Label1.Top
End Sub

Private Sub Label2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Source.Left = Label2.Left  'Locks the control into position over Label2
Source.Top = Label2.Top
End Sub

Private Sub Label3_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Source.Left = Label3.Left  'Locks the control into position over Label3
Source.Top = Label3.Top
End Sub
