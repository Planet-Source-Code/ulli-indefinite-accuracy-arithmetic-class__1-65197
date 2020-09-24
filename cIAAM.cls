VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIAAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'____________________________________
'Indefinite Accuracy Arithmetic Class  © 2003 UMGEDV GmbH
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'This class contains the basic arithmetic operations (+ - * / mod) and a square root function for
'integer operands of virtually any size. A few helper functions are also included.
'
'All values are stored in zero based arrays of bytes with leading LSD (little endian).
'The sign is stored in any or all of the upper four bits of the LSD.
'All values are treated as integers; if you need to work with fractions you will have to scale by powers of 10
'using the Shift Function to get the accuracy you require. Results are always truncated towards zero.
'
'Examples
'           7 / 4 = 1 r3
'
'           70 / 4 = 17 r2
'
'           700 / 4 = 175 r0 --> also 7.00 / 4 = 1.75 or 70.0 / 4 = 17.5
'
'There is no real limit on the number of digits but your memory size and time (up to a theoretical maximum
'of more than four thousand million digits).
'
'However, extracting for example the root of 10,000-digit number giving a result of 5,000 accurate(!!) digits
'will let you drink a cup of coffee while waiting (faster when compiled -- about 27 seconds on my 1800 Athlon).
'_________________________________________________________________________________________________
'This class is mainly intended for cryptography where floating point rounding cannot be tolerated.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Option Explicit

Private Operand1()          As Byte
Private Operand2()          As Byte
Private tmp1()              As Byte
Private tmp2()              As Byte

Private Const SignBits      As Byte = &HF0
Private Const NumericBits   As Byte = Not SignBits

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function Add(ByRef Summand1() As Byte, ByRef Summand2() As Byte, ByRef Sum() As Byte) As Long

  'Adds Summand1 and Summand2 giving the Sum.
  'Returns the number of digits of the sum.

  Dim Carry     As Long 'carry and intermediate sum
  Dim i         As Long
  Dim SumLen    As Long
  Dim l1        As Long
  Dim l2        As Long
  Dim Sign      As Byte

    Select Case True

      Case Summand1(0) > 9 And Summand2(0) < 10     '1st neg and 2nd pos
        Summand1(0) = Summand1(0) And NumericBits   'so remove the sign and
        Add = Subtract(Summand2, Summand1, Sum)     'subtract 1st from 2nd
        Summand1(0) = Summand1(0) Or SignBits       'restore sign

      Case Summand1(0) < 10 And Summand2(0) > 9     '1st pos and 2nd neg
        Summand2(0) = Summand2(0) And NumericBits
        Add = Subtract(Summand1, Summand2, Sum)     'so subtract 2nd from 1st
        Summand2(0) = Summand2(0) Or SignBits

      Case Else                                     'same sign
        Copy Summand1, Operand1
        Copy Summand2, Operand2
        Sign = Operand1(0) And SignBits
        Operand1(0) = Operand1(0) And NumericBits
        Operand2(0) = Operand2(0) And NumericBits
        l1 = RemoveLeadingZerosFrom(Operand1)
        l2 = RemoveLeadingZerosFrom(Operand2)
        SumLen = Max(l1, l2)
        l1 = Min(l1, l2)
        ReDim Sum(0 To SumLen + 1)

        For i = 0 To SumLen
            Select Case True
              Case i <= l1
                Carry = Operand1(i) + Operand2(i) + Carry
              Case SumLen = l2
                Carry = Operand2(i) + Carry
              Case Else
                Carry = Operand1(i) + Carry
            End Select
            Select Case Carry
              Case 0
                'do nothing
              Case Is < 10
                Sum(i) = Carry
                Carry = 0
              Case Else
                Sum(i) = Carry - 10
                Carry = 1
            End Select
        Next i

        If Carry Then
            Sum(i) = 1
        End If
        Sum(0) = Sum(0) Or Sign
        Add = RemoveLeadingZerosFrom(Sum) + 1
    End Select
    DoEvents

End Function

Private Sub Class_Initialize()

    If InIDE Then
        MsgBox "Please compile me; I am faster when compiled.", , "Indefinite Accuracy Arithmetic Class"
    End If

End Sub

Public Function Compare(ByRef FirstOperand() As Byte, ByRef SecondOperand() As Byte) As Long

  'Arithmetically compares two numbers.
  'Returns 1 if 1st is greater,  -1 if 2nd is greater, and 0 if equal.

  Dim l1    As Long
  Dim l2    As Long

    Copy FirstOperand, Operand1
    Copy SecondOperand, Operand2
    l1 = RemoveLeadingZerosFrom(Operand1)
    l2 = RemoveLeadingZerosFrom(Operand2)
    Select Case True

      Case (Operand1(0) And SignBits) = 0 And (Operand2(0) And SignBits) <> 0 '1st pos and 2nd neg
        Compare = 1

      Case (Operand1(0) And SignBits) <> 0 And (Operand2(0) And SignBits) = 0 '1st neg and 2nd pos
        Compare = -1

      Case l1 > l2 '1st longer than 2nd
        Compare = 1

      Case l2 > l1 '2nd longer thwn 1st
        Compare = -1

      Case Else 'same number of digits and same sign
        Operand1(0) = Operand1(0) And NumericBits
        Operand2(0) = Operand2(0) And NumericBits

        For l2 = l1 To 0 Step -1
            Select Case Operand1(l2)
              Case Is > Operand2(l2)
                Compare = 1
                Exit For 'loop varying l2
              Case Is < Operand2(l2)
                Compare = -1
                Exit For 'loop varying l2
            End Select
        Next l2

    End Select

    If (FirstOperand(0) And SignBits) <> 0 And (SecondOperand(0) And SignBits) <> 0 Then 'both neg
        Compare = -Compare
    End If

    DoEvents

End Function

Public Sub Copy(ByRef Source() As Byte, ByRef Destination() As Byte)

  'Copies a number.

  Dim i As Long

    i = UBound(Source)
    ReDim Destination(0 To i)
    CopyMemory Destination(0), Source(0), i + 1

End Sub

Public Function Create(ByVal Value As Long) As Byte()

  'Creates a number form a long, returns the created number.

  Dim i     As Long
  Dim Sign  As Byte

    Sign = Sgn(Value) And SignBits
    Value = Abs(Value)
    ReDim Operand1(0 To 10)
    Do
        Operand1(i) = Value Mod 10
        Value = Value \ 10
        i = i + 1
    Loop Until Value = 0
    RemoveLeadingZerosFrom Operand1
    Operand1(0) = Operand1(0) Or Sign
    Create = Operand1

End Function

Public Function Divide(ByRef Dividend() As Byte, ByRef Divisor() As Byte, ByRef Quotient() As Byte, ByRef Remainder() As Byte) As Long

  'Divides Dividend by Divisor giving Quotient and Remainder such that
  'quotient * divisor + remainder = dividend; the remainder has the same sign as the dividend so that
  'the quotient is always truncated towards zero.

  'examples:
  '            7 / 3  = 2 r1
  '            7 / -3 = -2 r1
  '           -7 / 3  = -2 r-1
  '           -7 / -3 = 2 r-1

  'If quotiend and remainder have same memory address then the remainder is discarded.

  'Returns the number of digits of the quotient.

  Dim e             As Long
  Dim Borrow        As Long 'borrow and intermediate difference
  Dim i             As Long
  Dim j             As Long
  Dim ij            As Long
  Dim DvdLen        As Long
  Dim DvrLen        As Long
  Dim DvrLenPlus1   As Long
  Dim SignQ         As Byte
  Dim SignR         As Byte

    Copy Dividend, Operand1
    Copy Divisor, Operand2
    ReDim Remainder(0 To 0)
    If (Operand1(0) > 9) Xor (Operand2(0) > 9) Then 'different sign
        SignQ = SignBits
        SignR = Operand1(0) And SignBits
    End If
    Operand1(0) = Operand1(0) And NumericBits
    Operand2(0) = Operand2(0) And NumericBits
    DvdLen = RemoveLeadingZerosFrom(Operand1)
    DvrLen = RemoveLeadingZerosFrom(Operand2)
    If DvdLen < 0 Then
        ReDim Quotient(0 To 0)
        Divide = 1
      Else 'NOT DvdLen...
        DvrLenPlus1 = DvrLen + 1
        ReDim Preserve Operand2(0 To DvrLenPlus1)
        If DvrLen >= 0 Then
            If DvdLen < DvrLen Then
                ReDim Quotient(0 To 0)
              Else 'NOT DvdLen...
                ReDim Quotient(0 To DvdLen - DvrLen + 1)
                If DvrLen <= DvdLen Then
                    For i = DvdLen - DvrLen To 0 Step -1
                        Quotient(i) = 0
                        Borrow = 0
                        Do
                            For j = 0 To DvrLen
                                ij = i + j
                                Borrow = CLng(Operand1(ij)) - Operand2(j) - Borrow
                                If Borrow < 0 Then
                                    Operand1(ij) = Borrow + 10
                                    Borrow = 1
                                  Else 'NOT BORROW...
                                    Operand1(ij) = Borrow
                                    Borrow = 0
                                End If
                            Next j
                            On Error Resume Next
                                Quotient(i) = Quotient(i) + 1
                                e = Err
                            On Error GoTo 0
                            If e Then
                                Err.Raise 11 'division by zero
                            End If
                        Loop While Borrow = 0
                        Quotient(i) = Quotient(i) - 1
                        'add back once
                        Borrow = 0
                        For j = 0 To DvrLen
                            ij = i + j
                            Borrow = Operand1(ij) + Operand2(j) + Borrow
                            If Borrow < 10 Then
                                Operand1(ij) = Borrow
                                Borrow = 0
                              Else 'NOT BORROW...
                                Operand1(ij) = Borrow - 10
                                Borrow = 1
                            End If
                        Next j
                        DvrLen = DvrLenPlus1
                        DoEvents
                    Next i
                End If
            End If
            Divide = RemoveLeadingZerosFrom(Quotient) + 1
            Quotient(0) = Quotient(0) Or SignQ
            If VarPtr(Quotient(0)) <> VarPtr(Remainder(0)) Then
                RemoveLeadingZerosFrom Operand1
                Copy Operand1, Remainder
                Remainder(0) = Remainder(0) Or SignR
            End If
        End If
    End If

End Function

Private Function InIDE(Optional c As Boolean = False) As Boolean

  Static b As Boolean

    b = c
    If b = False Then
        Debug.Assert InIDE(True)
    End If
    InIDE = b

End Function

Private Function Max(ByVal a As Long, ByVal b As Long) As Long

    If a >= b Then
        Max = a
      Else 'NOT A...
        Max = b
    End If

End Function

Private Function Min(ByVal a As Long, ByVal b As Long) As Long

    If a <= b Then
        Min = a
      Else 'NOT A...
        Min = b
    End If

End Function

Public Function Multiply(ByRef Factor1() As Byte, ByRef Factor2() As Byte, ByRef Product() As Byte) As Long

  'Multiplies Factor1 by Factor2 giving Product.
  'Returns the number of digits of the product.

  Dim Carry     As Long 'carry and intermediate product
  Dim Leng      As Long
  Dim i         As Long
  Dim j         As Long
  Dim ij        As Long
  Dim Fac1Len   As Long
  Dim Fac2Len   As Long
  Dim Sign      As Byte

    Copy Factor1, Operand1
    Copy Factor2, Operand2
    If (Operand1(0) > 9) Xor (Operand2(0) > 9) Then
        Sign = SignBits
    End If
    Operand1(0) = Operand1(0) And NumericBits
    Operand2(0) = Operand2(0) And NumericBits
    Fac1Len = RemoveLeadingZerosFrom(Operand1)
    Fac2Len = RemoveLeadingZerosFrom(Operand2)
    ReDim Product(Fac1Len + Fac2Len + 1)
    For i = 0 To Fac1Len
        Carry = 0
        For j = 0 To Fac2Len
            ij = i + j
            Carry = Product(ij) + Operand1(i) * Operand2(j) + Carry
            Select Case Carry
              Case 0
                'do nothing
              Case Is > 9
                Product(ij) = Carry Mod 10
                Carry = Carry \ 10
                Leng = ij
              Case Else
                Product(ij) = Carry
                Carry = 0
                Leng = ij
            End Select
        Next j
        If Carry Then
            Leng = ij + 1
            Product(Leng) = Carry
        End If
        DoEvents
    Next i
    ReDim Preserve Product(0 To Leng)
    Product(0) = Product(0) Or Sign
    Multiply = RemoveLeadingZerosFrom(Product) + 1

End Function

Private Function RemoveLeadingZerosFrom(ByRef Number() As Byte) As Long

    For RemoveLeadingZerosFrom = UBound(Number) To 0 Step -1
        If Number(RemoveLeadingZerosFrom) Then
            Exit For 'loop varying removeleadingzerosfrom
        End If
    Next RemoveLeadingZerosFrom
    If RemoveLeadingZerosFrom >= 0 Then
        ReDim Preserve Number(0 To RemoveLeadingZerosFrom)
      Else 'NOT REMOVELEADINGZEROSFROM...
        ReDim Number(0 To 0)
        RemoveLeadingZerosFrom = 0
    End If
    DoEvents

End Function

Public Function Shift(ByRef Number() As Byte, ByVal Places As Long) As Long

  'Scales Number up or down by 10 ^ places  (no rounding)
  '   Shift 45678, -1 becomes 4567
  '   Shift 45678, +2 becomes 4567800

  'Returns the number of places of the result.

  Dim i     As Long
  Dim Sign  As Byte

    Sign = Number(0) And SignBits
    Number(0) = Number(0) And NumericBits
    Select Case Places
      Case Is < 0
        ReDim Preserve Number(UBound(Number) - Places)
        For i = -Places To UBound(Number)
            Number(i + Places) = Number(i)
        Next i
      Case Is > 0
        ReDim Preserve Number(UBound(Number) + Places)
        For i = UBound(Number) - Places To 0 Step -1
            Number(i + Places) = Number(i)
        Next i
        For i = 0 To Places - 1
            Number(i) = 0
        Next i
    End Select
    Number(0) = Number(0) Or Sign
    Shift = RemoveLeadingZerosFrom(Number) + 1
    DoEvents

End Function

Public Function SquareRoot(ByRef Radicand() As Byte, ByRef Root() As Byte) As Long

  'Computes the square root of Radicand and places it in Root; the Radicand may be scaled up (shifted left).
  'Returns the number of digits of the root,

  'You must fill Root with an estimate of the resulting root using the appropriate number of digits
  'before calling this function. Estimate can be anything but zero -- a good estimate however
  'will speed up things. Negative estimates return the negative root.

  'Uses the Babylonian or Heron's Algorithm (see http://www.mathpages.com/home/kmath190.htm)

  'X[n+1] = (X[n] + R / X[n]) / 2

  'This algo converges quadratically; so with a decent estimate approx log2(num_digits_of_root) iterations
  'are required, plus one extra to verify that there was no improvement anymore.

    If RemoveLeadingZerosFrom(Radicand) = 0 And (Radicand(0) And NumericBits) = 0 Then
        ReDim Root(0 To 0)
      Else 'NOT (RADICAND(0)... 'NOT REMOVELEADINGZEROSFROM(RADICAND)...
        If Radicand(0) And SignBits Then
            Err.Raise 5, , "Radicand cannot be negative"
          Else 'NOT RADICAND(0)...
            Shift Radicand, 2 * RemoveLeadingZerosFrom(Root) 'scale radicand
            Do
                Copy Root, tmp2                     'save current aproximation as X(n)

                Divide Radicand, Root, tmp1, tmp1   'R / X(n) -> tmp
                Add Root, tmp1, tmp1                'X(n) + R / X(n) -> tmp
                Divide tmp1, Create(2), Root, Root  '(X(n) + R / X(n)) / 2 -> new approximation X(n+1)

            Loop Until Compare(Root, tmp2) = 0      'until no improvement
        End If
    End If
    SquareRoot = RemoveLeadingZerosFrom(Root) + 1

End Function

Public Function Subtract(ByRef Minuend() As Byte, ByRef Subtrahend() As Byte, ByRef Difference() As Byte) As Long

  'Subtracts Subtrahend from Minuend giving the Difference.
  'Returns the number of digits of the difference.

  Dim Borrow    As Long 'borrow and intermediate difference
  Dim i         As Long
  Dim DifLen    As Long
  Dim l1        As Long
  Dim l2        As Long

    Select Case True
      Case Minuend(0) < 10 And Subtrahend(0) > 9                '1st pos and 2nd neg
        Subtrahend(0) = Subtrahend(0) And NumericBits
        Subtract = Add(Minuend, Subtrahend, Difference)         'so add them
        Subtrahend(0) = Subtrahend(0) Or SignBits

      Case Minuend(0) > 9 And Subtrahend(0) < 10                '1st neg and 2nd pos
        Minuend(0) = Minuend(0) And NumericBits
        Subtract = Add(Minuend, Subtrahend, Difference)         'so add them
        Difference(0) = Difference(0) Or SignBits               'and make diff neg
        Minuend(0) = Minuend(0) Or SignBits

      Case Minuend(0) > 9 And Subtrahend(0) > 9                 'both neg
        Minuend(0) = Minuend(0) And NumericBits
        Subtrahend(0) = Subtrahend(0) And NumericBits
        Subtract = Subtract(Subtrahend, Minuend, Difference)    'so subtract in reverse
        Minuend(0) = Minuend(0) Or SignBits
        Subtrahend(0) = Subtrahend(0) Or SignBits

      Case Else                                                 'both pos
        Copy Minuend, Operand1
        Copy Subtrahend, Operand2
        l1 = RemoveLeadingZerosFrom(Operand1)
        l2 = RemoveLeadingZerosFrom(Operand2)
        If l1 >= l2 Then
            DifLen = l1
            l1 = Min(l1, l2)
            ReDim Difference(0 To DifLen + 1)
            For i = 0 To DifLen
                If i <= l1 Then
                    Borrow = CLng(Operand1(i)) - Operand2(i) - Borrow
                  Else 'NOT I...
                    Borrow = CLng(Operand1(i)) - Borrow
                End If
                Select Case Borrow
                  Case 0
                    'do nothing
                  Case Is > 0
                    Difference(i) = Borrow
                    Borrow = 0
                  Case Else
                    Difference(i) = Borrow + 10
                    Borrow = 1
                End Select
            Next i
            If Borrow Then 'underflow
                Subtract = Subtract(Subtrahend, Minuend, Difference) 'do it again in reverse
                Difference(0) = Difference(0) Or SignBits 'it's faster to take chances than to compare every time
            End If
          Else 'NOT l1...
            Subtract = Subtract(Subtrahend, Minuend, Difference)
            Difference(0) = Difference(0) Or SignBits
        End If
    End Select
    Subtract = RemoveLeadingZerosFrom(Difference) + 1
    DoEvents

End Function

':) Ulli's VB Code Formatter V2.21.6 (2006-Mai-03 00:13)  Decl: 37  Code: 510  Total: 547 Lines
':) CommentOnly: 63 (11,5%)  Commented: 48 (8,8%)  Empty: 90 (16,5%)  Max Logic Depth: 9
