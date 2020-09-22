Attribute VB_Name = "modBitset"
' ============================------.,
'
' BIT-SET Operations in Visual Basic Module
' ====================================================--.,
' Function Set to Handle upto 128 Bits of Data
'
' written by: Chris Vega [gwapo@models.com]
'
' =====================================================================---.,

Global Const Author = "Chris Vega [gwapo@models.com]"
Global Const CopyRight = "(C) Copyright 2001 by " & Author

' ===========================================================---.,
' int2bin
'
' Unsigned Conversion of Long Integer Value to Binary String
'
' INPUT:
'
'   numArgument = Long Number Value to convert to Binary String
'
'                 Max = 18,446,744,073,709,551,615
'                 Min = 0
'
' =====================================================================---.,
Function int2bin(numArgument As Variant) As String
    int2bin = ""
    While numArgument > 0
        If Int(numArgument / 2) < (numArgument / 2) _
            Then int2bin = "1" & int2bin _
            Else int2bin = "0" & int2bin
        numArgument = Int(numArgument / 2)
    Wend
    
    If Len(int2bin) = 0 Then int2bin = 0
End Function

' ===========================================================---.,
' bin2int
'
' Unsigned Conversion of Binary String to Long Integer Value
'
' INPUT:
'
'   binArgument = Binary String to convert to Long Number
'
'                 Max = 18,446,744,073,709,551,615
'                 Min = 0
'
' =====================================================================---.,
Function bin2int(binArgument As Variant) As Variant
    Dim binValue As Variant
    
    binValue = 1
    bin2int = 0
    
    For i = Len(binArgument) To 1 Step -1
        If longval(Mid(binArgument, i, 1)) Then _
            bin2int = bin2int + binValue
        binValue = binValue * 2
    Next
End Function

' ===========================================================---.,
' longval
'
' Converts String to a Long Number Value
'
' The val() function of Visual Basic only supports 32-bits of Data,
' therefore, longval() function is used among all functions here in
' place of val() to convert String to Number.
'
' INPUT:
'
'   numArgument = String to Convert to Number
'
'                 Max = 18,446,744,073,709,551,615
'                 Min = 0
'
' =====================================================================---.,
Function longval(ByVal numArgument As Variant) As Variant
    Dim xvalue As Variant
    longval = 0
    xvalue = 1
    For i = Len(numArgument) To 1 Step -1
        If ((Mid(numArgument, i, 1) >= "0") And _
            (Mid(numArgument, i, 1) <= "9")) Then
            longval = longval + (Val(Mid(numArgument, i, 1)) * xvalue)
            xvalue = xvalue * 10
        End If
    Next
End Function

' ===========================================================---.,
' shl
'
' Shift Number Argument to Left, ie, Removes Bit(s) from the Right
'
' INPUT:
'
'   numArgument = Number to Shift
'
'                 Max = 18,446,744,073,709,551,615
'                 Min = 0
'
'   countOfShift = How many to Shift, must be valid, error checkings must
'                  be done before the call to this function.
'
' =====================================================================---.,
Function shl(ByVal numArgument As Variant, _
             Optional ByVal countOfShift As Long = 1) As Variant
    If longval(numArgument) > 0 Then
        numArgument = int2bin(numArgument)
        
        shl = bin2int(Left(numArgument, _
                       Len(numArgument) - countOfShift))
    Else
        shl = 0
    End If
End Function

' ===========================================================---.,
' shr
'
' Shift Number Argument to Right, ie, Removes Bit(s) from the Left
'
' INPUT:
'
'   numArgument = Number to Rotate
'
'                 Max = 18,446,744,073,709,551,615
'                 Min = 0
'
'   countOfShift = How many to Shift, must be valid, error checkings must
'                  be done before the call to this function.
'
' =====================================================================---.,
Function shr(ByVal numArgument As Variant, _
             Optional ByVal countOfShift As Long = 1) As Variant
    If longval(numArgument) > 0 Then _
        shr = bin2int(int2bin(numArgument) & String(countOfShift, 48)) _
    Else _
        shr = 0
End Function

' ===========================================================---.,
' rol
'
' Rotate Bit of Number Argument to Left, ie, Moves Bit(s) from the Right
'
' INPUT:
'
'   numArgument = Number to Rotate
'
'                 Max = 18,446,744,073,709,551,615
'                 Min = 0
'
'   countOfRotate = How many Bits to Rotate, must be valid, error checkings
'                   must be done before the call to this function.
'
' =====================================================================---.,
Function rol(ByVal numArgument As Variant, _
             Optional ByVal countOfRotate As Long = 1) As Variant
    If longval(numArgument) > 0 Then
        numArgument = int2bin(numArgument)
        
        rol = bin2int( _
                      Right(numArgument, countOfRotate) & _
                       Left(numArgument, _
                        Len(numArgument) - countOfRotate))
    Else
        rol = 0
    End If
End Function

' ===========================================================---.,
' ror
'
' Rotate Bit of Number Argument to Right, ie, Moves Bit(s) from the Right
'
' INPUT:
'
'   numArgument = Number to Shift
'
'                 Max = 18,446,744,073,709,551,615
'                 Min = 0
'
'   countOfRotate = How many Bits to Rotate, must be valid, error checkings
'                   must be done before the call to this function.
'
' =====================================================================---.,
Function ror(ByVal numArgument As Variant, _
             Optional ByVal countOfRotate As Long = 1) As Variant
    If longval(numArgument) > 0 Then
        numArgument = int2bin(numArgument)
        
        ror = bin2int( _
                       Right(numArgument, _
                         Len(numArgument) - countOfRotate) & _
                        Left(numArgument, countOfRotate))
    Else
        ror = 0
    End If
End Function

' ===========================================================---.,
' not_
'
' Unsigned NOT Operator, or negate the Bits
'
' INPUT:
'
'   numArgument = Number to Negate
'
'                 Max = 18,446,744,073,709,551,615
'                 Min = 0
'
' =====================================================================---.,
Function not_(ByVal numArgument As Variant) As Variant
    If longval(numArgument) > 0 Then
        Dim i As Long
        numArgument = int2bin(numArgument)
        
        not_ = ""
        For i = 1 To Len(numArgument)
            If Asc(Mid(numArgument, i, 1)) - 48 _
               Then not_ = not_ & "0" _
               Else not_ = not_ & "1"
        Next
        not_ = bin2int(not_)
    End If
End Function
