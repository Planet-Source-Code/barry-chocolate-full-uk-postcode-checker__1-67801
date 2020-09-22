VERSION 5.00
Begin VB.Form frmPostcodeChecker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Postcode Checker"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtPostcode 
      Height          =   285
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblPostcodeChecker 
      Caption         =   "Please enter postode"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmPostcodeChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheck_Click()
    On Error GoTo errCheck
    'calls the function to check if the postcode is valid
    'displays a messge box with the result
    'then clears the postcode box
    If IsPostcodeValid(txtPostcode.Text) = True Then
        MsgBox "The postcode " & txtPostcode & " is valid", vbInformation + vbOKOnly, "Valid!"
        txtPostcode.Text = ""
    Else
        MsgBox "The postcode " & txtPostcode & " is invalid", vbInformation + vbOKOnly, "Invalid!"
        txtPostcode.Text = ""
    End If
    Exit Sub
errCheck:
    'displays a message box with the error number and description
    MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Private Sub txtPostcode_Validate(Cancel As Boolean)
    'Changes the case of the postcode to upper case
    txtPostcode.Text = UCase(txtPostcode.Text)
End Sub



