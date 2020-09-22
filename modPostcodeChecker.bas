Attribute VB_Name = "modPostcodeChecker"
Option Explicit

'checks whether a Postcode is in one of the 6 UK formats
'Returns true if the postcode is in the correct format
'The postcode should be passed as A string with a space seperating
'the Outcode(first part of the postcode) from the Incode (end part of the postcode)

'Valid Formats (A = Alpha, N = Nueric)
'Format     Example
'AN NAA     M1 1AA
'AAN NAA    CR2 6XH
'ANA NAA    W1A 1HQ
'ANN NAA    M60 1NW
'AANA NAA   EC1A 1BB
'AANN NAA   DN55 1PT

'The letters Q, V and X are not used in the first position.
'The letters I, J and Z are not used in the second position.
'The only letters to appear in the third position are A, B, C, D, E, F, G, H, J, K, S, T, U and W.
'The only letters to appear in the fourth position are A, B, E, H, M, N, P, R, V, W, X and Y.
'The second half of the Postcode is always consistent numeric, alpha, alpha format and the letters C, I, K, M, O and V are never used.
'*GIR 0AA is a Postcode that was issued historically and does not confirm to current rules on valid Postcode formats, It is however, still in use.

Public Function IsPostcodeValid(ByRef Postcode As String) As Boolean
    On Error GoTo ErrIsPostcodeValid
    IsPostcodeValid = False
    'Checks the postcode is a valid length
    If Len(Postcode) < 6 Or Len(Postcode) > 8 Then
        Exit Function
    End If
    'Changes the postcode to upper case to make checking using like easier
    Postcode = UCase(Postcode)
    'AN NAA
    If Postcode Like "[A-PR-UWY-Z]# #[A-BD-MJLNP-UW-Z][A-BD-MJLNP-UW-Z]" Then
        IsPostcodeValid = True
        Exit Function
    End If
    'AAN NAA
    If Postcode Like "[A-PR-UWY-Z][A-HK-Y]# #[A-BD-MJLNP-UW-Z][A-BD-MJLNP-UW-Z]" Then
        IsPostcodeValid = True
        Exit Function
    End If
    'ANA NAA
    If Postcode Like "[A-PR-UWY-Z]#[A-HJ-KS-UW] #[A-BD-MJLNP-UW-Z][A-BD-MJLNP-UW-Z]" Then
        IsPostcodeValid = True
        Exit Function
    End If
    'ANN NAA
    If Postcode Like "[A-PR-UWY-Z]## #[A-BD-MJLNP-UW-Z][A-BD-MJLNP-UW-Z]" Then
        IsPostcodeValid = True
        Exit Function
    End If
    'AANA NAA
    If Postcode Like "[A-PR-UWY-Z][A-HK-Y]#[A-BEHM-NPRV-Y] #[A-BD-MJLNP-UW-Z][A-BD-MJLNP-UW-Z]" Then
        IsPostcodeValid = True
        Exit Function
    End If
    'AANN NAA
    If Postcode Like "[A-PR-UWY-Z][A-HK-Y]## #[A-BD-MJLNP-UW-Z][A-BD-MJLNP-UW-Z]" Then
        IsPostcodeValid = True
        Exit Function
    End If
    'GIR 0AA is a postcode that was issued historically and does not confirm to current rules on valid postcode formats
    'it is however, still in use
    If Postcode = "GIR 0AA" Then
        IsPostcodeValid = True
        Exit Function
    End If
    Exit Function
ErrIsPostcodeValid:
    IsPostcodeValid = False
    MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbOKOnly, "Postcode Check Error"
End Function

