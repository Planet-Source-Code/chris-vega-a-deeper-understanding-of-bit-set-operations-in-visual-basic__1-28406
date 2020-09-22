VERSION 5.00
Begin VB.Form frmGetNumber 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3060
      TabIndex        =   2
      Top             =   900
      Width           =   1410
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   4545
      TabIndex        =   1
      Top             =   900
      Width           =   1410
   End
   Begin VB.TextBox txtNumber 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      MaxLength       =   16
      TabIndex        =   0
      Top             =   495
      Width           =   5880
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   135
      Width           =   45
   End
End
Attribute VB_Name = "frmGetNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ReturnValue As Variant
Public MaxValue As Variant

Private Sub cmdCancel_Click()
    ReturnValue = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If frmBitSetOperations.optNumber(0).Value Then
        ReturnValue = longval(txtNumber)
    Else
        ReturnValue = bin2int(txtNumber)
    End If
    
    If ReturnValue <= MaxValue Then _
        Unload Me Else _
        MsgBox "Maximum Value is : " & MaxValue
End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 _
       And KeyAscii <= frmBitSetOperations.maxLimit) _
       And (KeyAscii <> 8) _
       And (KeyAscii <> 27) Then
        SendKeys "{bs}"
    End If
End Sub
