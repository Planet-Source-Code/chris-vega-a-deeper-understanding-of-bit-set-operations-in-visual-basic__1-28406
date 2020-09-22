VERSION 5.00
Begin VB.Form frmBitSetOperations 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bit-Set Demo In Visual Basic"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   Icon            =   "BitSet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optNumber 
      Caption         =   "Binary"
      Height          =   195
      Index           =   1
      Left            =   270
      TabIndex        =   13
      Top             =   1395
      Width           =   1590
   End
   Begin VB.OptionButton optNumber 
      Caption         =   "Long Integer"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   12
      Top             =   810
      Value           =   -1  'True
      Width           =   1545
   End
   Begin VB.CommandButton cmddone 
      Caption         =   "done"
      Height          =   375
      Left            =   4950
      TabIndex        =   11
      Top             =   1755
      Width           =   975
   End
   Begin VB.CommandButton cmdeqv 
      Caption         =   "eqv"
      Height          =   375
      Left            =   4950
      TabIndex        =   10
      Top             =   945
      Width           =   975
   End
   Begin VB.CommandButton cmdimp 
      Caption         =   "imp"
      Height          =   375
      Left            =   4950
      TabIndex        =   9
      Top             =   540
      Width           =   975
   End
   Begin VB.CommandButton cmdxor 
      Caption         =   "xor"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   1755
      Width           =   975
   End
   Begin VB.CommandButton cmdor 
      Caption         =   "or"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   1350
      Width           =   975
   End
   Begin VB.CommandButton cmdand 
      Caption         =   "and"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   945
      Width           =   975
   End
   Begin VB.CommandButton cmdnot 
      Caption         =   "not"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   540
      Width           =   975
   End
   Begin VB.CommandButton cmdror 
      Caption         =   "ror"
      Height          =   375
      Left            =   2205
      TabIndex        =   4
      Top             =   1755
      Width           =   975
   End
   Begin VB.CommandButton cmdrol 
      Caption         =   "rol"
      Height          =   375
      Left            =   2205
      TabIndex        =   3
      Top             =   1350
      Width           =   975
   End
   Begin VB.CommandButton cmdshl 
      Caption         =   "shl"
      Height          =   375
      Left            =   2205
      TabIndex        =   2
      Top             =   540
      Width           =   975
   End
   Begin VB.CommandButton cmdshr 
      Caption         =   "shr"
      Height          =   375
      Left            =   2205
      TabIndex        =   1
      Top             =   945
      Width           =   975
   End
   Begin VB.TextBox txtNumber 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Top             =   90
      Width           =   5880
   End
   Begin VB.Shape shShadow 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   285
      Left            =   135
      Top             =   135
      Width           =   5865
   End
   Begin VB.Label lblStatbar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   45
      TabIndex        =   14
      Top             =   2205
      Width           =   5910
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   240
      Left            =   90
      Top             =   2250
      Width           =   5910
   End
End
Attribute VB_Name = "frmBitSetOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public maxLimit As Long
Public BinaryOn As Boolean

Private Sub cmdand_Click()
    Dim andValue As Long
    
    andValue = GetNumber("Enter Number to be use in And", 65535)

    lblStatbar.Caption = "and " & TNumber & ", " & andValue
    txtNumber = WNumber(TNumber And andValue)
    txtNumber.SetFocus
End Sub

Private Sub cmddone_Click()
    Unload Me
End Sub

Private Sub cmdeqv_Click()
    Dim eqvValue As Long
    
    eqvValue = GetNumber("Enter Number to be use in Equvalent", 65535)

    lblStatbar.Caption = "eqv " & TNumber & ", " & eqvValue
    txtNumber = WNumber(TNumber Eqv eqvValue)
    txtNumber.SetFocus
End Sub

Private Sub cmdimp_Click()
    Dim impValue As Long
    
    impValue = GetNumber("Enter Number to be use in Impression", 65535)

    lblStatbar.Caption = "imp " & TNumber & ", " & impValue
    txtNumber = WNumber(TNumber Imp impValue)
    txtNumber.SetFocus
End Sub

Private Sub cmdnot_Click()
    '
    ' To pass the Signed NOT operator of Visual Basic, i used another
    ' function, i called it not_()
    '
    lblStatbar = "not " & TNumber
    txtNumber = WNumber(not_(TNumber))
    txtNumber.SetFocus
End Sub

Private Sub cmdor_Click()
    Dim orValue As Long
    
    orValue = GetNumber("Enter Number to be use in Or", 65535)

    lblStatbar.Caption = "or " & TNumber & ", " & orValue
    txtNumber = WNumber(TNumber Or orValue)
    txtNumber.SetFocus
End Sub

Private Sub cmdrol_Click()
    Dim rolValue As Long
    
    rolValue = GetNumber("Enter Number to be use in Rotating", 8)

    If Len(int2bin(TNumber)) > Len(int2bin(rolValue)) Then
        lblStatbar.Caption = "rol " & TNumber & ", " & rolValue
        txtNumber = WNumber(rol(TNumber, rolValue))
    Else
        lblStatbar.Caption = "Illegal Operations"
    End If
    
    txtNumber.SetFocus
End Sub

Private Sub cmdror_Click()
    Dim rorValue As Long
    
    rorValue = GetNumber("Enter Number to be use in Rotating", 8)

    If Len(int2bin(TNumber)) > Len(int2bin(rorValue)) Then
        lblStatbar.Caption = "ror " & TNumber & ", " & rorValue
        txtNumber = WNumber(ror(TNumber, rorValue))
    Else
        lblStatbar.Caption = "Illegal Operations"
    End If
    
    txtNumber.SetFocus
End Sub

Private Sub cmdshl_Click()
    Dim shlValue As Long
    
    shlValue = GetNumber("Enter Number to be use in Shifting", 8)

    If Len(int2bin(TNumber)) > Len(int2bin(shlValue)) Then
        lblStatbar.Caption = "shl " & TNumber & ", " & shlValue
        txtNumber = WNumber(shl(TNumber, shlValue))
    Else
        lblStatbar.Caption = "Illegal Operations"
    End If
    
    txtNumber.SetFocus
End Sub

Private Sub cmdshr_Click()
    Dim shrValue As Long
    
    shrValue = GetNumber("Enter Number to be use in Shifting", 8)

    lblStatbar.Caption = "shr " & TNumber & ", " & shrValue
    txtNumber = WNumber(shr(TNumber, shrValue))
    
    txtNumber.SetFocus
End Sub

Private Sub cmdxor_Click()
    Dim xorValue As Long
    
    xorValue = GetNumber("Enter Number to be use in XOr", 65535)

    lblStatbar.Caption = "xor " & TNumber & ", " & xorValue
    txtNumber = WNumber(TNumber Xor xorValue)
    txtNumber.SetFocus
End Sub

Private Sub Form_Load()
    maxLimit = 57
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox CopyRight, vbInformation, Caption
    End
End Sub

Private Sub optNumber_Click(Index As Integer)
    If Index Then
        maxLimit = 49
        txtNumber.MaxLength = 128
        BinaryOn = True
        If Len(txtNumber) > 0 Then txtNumber = int2bin(txtNumber)
    Else
        maxLimit = 57
        BinaryOn = False
        txtNumber.MaxLength = 16
        If Len(txtNumber) > 0 Then txtNumber = bin2int(txtNumber)
    End If
    
    txtNumber.SetFocus
End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 _
       And KeyAscii <= maxLimit) _
       And (KeyAscii <> 8) _
       And (KeyAscii <> 27) Then
        SendKeys "{bs}"
    End If
End Sub

Function GetNumber(lpMessage As String, MaxValue)
    With frmGetNumber
        .Caption = Caption
        .lblCaption = lpMessage
        .MaxValue = MaxValue
        .Show vbModal
        GetNumber = .ReturnValue
    End With
End Function

Function TNumber()
    If BinaryOn Then TNumber = bin2int(txtNumber) _
                Else TNumber = longval(txtNumber)
End Function

Function WNumber(numArgument As Variant)
    If BinaryOn Then WNumber = int2bin(numArgument) _
                Else WNumber = longval(numArgument)
End Function
