VERSION 5.00
Begin VB.Form frmAddress 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Test: cAddress Class"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDialog 
      Height          =   1335
      Index           =   0
      Left            =   1365
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmAddress.frx":0000
      Top             =   210
      Width           =   4530
   End
   Begin VB.Frame Frame1 
      Caption         =   "Address &Details:"
      Height          =   2850
      Left            =   105
      TabIndex        =   2
      Top             =   2625
      Width           =   5790
      Begin VB.TextBox txtDialog 
         Height          =   285
         Index           =   4
         Left            =   1395
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1680
         Width           =   4245
      End
      Begin VB.CheckBox chkDialog 
         Alignment       =   1  'Right Justify
         Caption         =   "&Auto Capitalise:"
         Height          =   330
         Left            =   105
         TabIndex        =   13
         Top             =   2415
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.ComboBox cboDialog 
         Height          =   315
         Index           =   0
         Left            =   1395
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   1995
         Width           =   4245
      End
      Begin VB.TextBox txtDialog 
         Height          =   285
         Index           =   3
         Left            =   1395
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1365
         Width           =   4245
      End
      Begin VB.TextBox txtDialog 
         Height          =   285
         Index           =   2
         Left            =   1395
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1050
         Width           =   4245
      End
      Begin VB.TextBox txtDialog 
         Height          =   705
         Index           =   1
         Left            =   1395
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "frmAddress.frx":0006
         Top             =   315
         Width           =   4245
      End
      Begin VB.Label lblDialog 
         Alignment       =   1  'Right Justify
         Caption         =   "C&ountry: "
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   11
         Top             =   2040
         Width           =   1170
      End
      Begin VB.Label lblDialog 
         Alignment       =   1  'Right Justify
         Caption         =   "&Postcode: "
         Height          =   225
         Index           =   5
         Left            =   135
         TabIndex        =   9
         Top             =   1710
         Width           =   1170
      End
      Begin VB.Label lblDialog 
         Alignment       =   1  'Right Justify
         Caption         =   "S&tate: "
         Height          =   225
         Index           =   4
         Left            =   135
         TabIndex        =   7
         Top             =   1395
         Width           =   1170
      End
      Begin VB.Label lblDialog 
         Alignment       =   1  'Right Justify
         Caption         =   "&City: "
         Height          =   225
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   1080
         Width           =   1170
      End
      Begin VB.Label lblDialog 
         Alignment       =   1  'Right Justify
         Caption         =   "&Street: "
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   315
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "&Close"
      Height          =   330
      Left            =   4725
      TabIndex        =   14
      Top             =   5670
      Width           =   1170
   End
   Begin VB.Label lblDialog 
      Alignment       =   1  'Right Justify
      Caption         =   "Format Layout:"
      Height          =   225
      Index           =   6
      Left            =   105
      TabIndex        =   16
      Top             =   1680
      Width           =   1170
   End
   Begin VB.Label lblInstructions 
      Caption         =   "Instructions"
      Height          =   855
      Left            =   1890
      TabIndex        =   15
      Top             =   1680
      Width           =   4005
   End
   Begin VB.Label lblDialog 
      Alignment       =   1  'Right Justify
      Caption         =   "&Address: "
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   1170
   End
End
Attribute VB_Name = "frmAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    frmAddress
' Author:       Slider
' Date:         03/08/2001
' Version:      01.00.00
' Description:  Test form for cAddress Class.
' Edit History: 01.00.00 03/08/01 Initial Release
'               01.00.01 04/08/01 Fixed Issue with cboDialog_Validate &
'                                 Auto-complete does not set ListIndex.
' Notes:        This test app was designed to exploit all of the features
'               available in the cAddress Class such as:-
'                   * Convert a formatted Full Address into individual fields
'                   * Convert individual fields to a formatted Full Address
'                   * Automatically capitalise fields/Full Address (option)
'                   * Validate the Full Address based on set criteria
'                   * Fills Combobox with all known countries
'               The test apps also illustrates (for beginners) how to:-
'                   * Auto-complete a ComboBox
'                   * Quick ComboBox search using API
'                   * Simple field hilighting methods for TextBox and
'                     ComboBox
'                   * Avoid complex If/Then structures using bitwise
'                     operation and the IIF function
'                   * Avoid infinite event loops (Stack overflow errors)
'                   * Encapsulating data and associated functions into a
'                     reusable code class
'
'===========================================================================

Option Explicit

Private Enum eTextBox
    etbAddress = 0
    etbStreet = 1
    etbCity = 2
    etbState = 3
    etbPostcode = 4
End Enum

Private Enum eComboBox
    ecbCountry = 0
End Enum

Private mcAddress   As cAddress

'Private mbCboLoading       As Boolean
Private mbCboExist(0 To 1) As Boolean
Private mbBackspaced       As Boolean
Private mbIsDirty          As Boolean
Private mbIsBusy           As Boolean

'/////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cboDialog_Change(Index As Integer)

    If mbIsBusy Then Exit Sub
'    If mbCboLoading Then Exit Sub

    mbIsDirty = True
    '## Set Search Toolbat Button state
    If Len(cboDialog(Index).Text) > 0 Then
        mbCboExist(Index) = True
    Else
        mbCboExist(Index) = False
    End If

    '## Auto-complete combobox
    '## If firing in response to a backspace or delete, don't run the auto-complete
    '   complete code. (Otherwise you wouldn't be able to back up.)
    If mbBackspaced = True Or cboDialog(Index).Text = "" Then
        mbBackspaced = False
        Exit Sub
    End If

    Dim lLoop As Long
    Dim nSel  As Long

    '## Run through the available items and grab the first matching one.
    For lLoop = 0 To cboDialog(Index).ListCount - 1
        If InStr(1, cboDialog(Index).List(lLoop), cboDialog(Index).Text, vbTextCompare) = 1 Then
            '## Save the SelStart property.
            nSel = cboDialog(Index).SelStart
            cboDialog(Index).Text = cboDialog(Index).List(lLoop)
            '## Set the selection in the combo.
            cboDialog(Index).SelStart = nSel
            cboDialog(Index).SelLength = Len(cboDialog(Index).Text) - nSel
            Exit For
        End If
    Next

End Sub

Private Sub cboDialog_Click(Index As Integer)
    If mbIsBusy Then Exit Sub
'    If mbCboLoading Then Exit Sub
    mbCboExist(Index) = True
    mbIsDirty = True            '## A change was made...
End Sub

Private Sub cboDialog_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    '## Auto-complete combobox
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If cboDialog(Index).Text <> "" Then
            '## Let the Change event know that it shouldn't respond to this change.
            mbBackspaced = True
        End If
    End If

End Sub

Private Sub cboDialog_KeyPress(Index As Integer, KeyAscii As Integer)
    Debug.Print KeyAscii
    If KeyAscii = 13 Then
        If mbCboExist(Index) Then
            '## Special code event...
        End If
    End If
End Sub

Private Sub cboDialog_Validate(Index As Integer, Cancel As Boolean)

    '## We're leaving this field...
    If mbIsDirty Then           '## Anything changed?
        mbIsBusy = True
        With cboDialog(Index)
            Select Case Index
                Case ecbCountry: mcAddress.Country = .Text
            End Select
            FindComboText cboDialog(Index), .Text
            txtDialog(etbAddress).Text = mcAddress.Address
        End With
        mbIsBusy = False
    End If
    mbIsDirty = False

End Sub

Private Sub chkDialog_Click()
    mcAddress.AutoCorrect = CBool(chkDialog.Value)
End Sub

Private Sub cmdDialog_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    lblInstructions = "Street" + vbCrLf + _
                      "Street..." + vbCrLf + _
                      "City, State Postcode" + vbCrLf + _
                      "Country"

    Set mcAddress = New cAddress
    mbIsBusy = True             '## disable the need to respond to specific VB events
    With mcAddress
        .AutoCorrect = CBool(chkDialog.Value)
        .Street = "Level 12, East Tower, Amp Place" + vbCrLf + "123 Vulture Street"
        .City = "North Ryde"
        .State = "New South Wales"
        .Postcode = "2123"
        .Country = "Australia"
        .FillComboBox cboDialog(ecbCountry), eactCountry
        '
        '## Fill GUI Fields with data
        '
        txtDialog(etbAddress).Text = .Address
        txtDialog(etbStreet).Text = .Street
        txtDialog(etbCity).Text = .City
        txtDialog(etbState).Text = .State
        txtDialog(etbPostcode).Text = .Postcode
        mbCboExist(ecbCountry) = FindComboText(cboDialog(ecbCountry), .Country)
    End With
    mbIsBusy = False            '## re-enable the app's VB event handling

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mcAddress = Nothing
End Sub

Private Sub txtDialog_Change(Index As Integer)
'    Debug.Print "txtDialog_Change"
    If mbIsBusy Then Exit Sub
    mbIsDirty = True            '## A change was made...
End Sub

Private Sub txtDialog_GotFocus(Index As Integer)
    HiLite txtDialog(Index)
End Sub

Private Sub txtDialog_Validate(Index As Integer, Cancel As Boolean)

'    Debug.Print "txtDialog_Validate"
    '## We're leaving this field...
    If mbIsDirty Then           '## Anything changed?
        mbIsBusy = True         '## disable the need to respond to specific VB events
        With mcAddress
            Select Case Index
                Case etbAddress
                    Dim eTest   As eAddressValidateFields
                    Dim eResult As eAddressValidateFields

                    .Address = txtDialog(etbAddress).Text
                    If CBool(chkDialog.Value) Then
                        txtDialog(etbAddress).Text = .Address
                    End If
                    txtDialog(etbStreet).Text = .Street     '## returns extracted fields
                    txtDialog(etbCity).Text = .City         '
                    txtDialog(etbState).Text = .State       '
                    txtDialog(etbPostcode).Text = .Postcode '
                    mbCboExist(ecbCountry) = FindComboText(cboDialog(ecbCountry), .Country)
                    '
                    '## Test if specific data was entered...
                    '
                    eTest = eavfStreet + eavfCity + eavfState + eavfPostcode
                    eResult = .ValidateAddress(txtDialog(etbAddress).Text, eTest)
                    If eResult <> eTest Then
                        MsgBox "Incomplete address. The following fields were missing:" + vbCrLf + vbCrLf + _
                                   IIf((eResult And eavfStreet), "", vbTab + "Street" + vbCrLf) + _
                                   IIf((eResult And eavfCity), "", vbTab + "City" + vbCrLf) + _
                                   IIf((eResult And eavfState), "", vbTab + "State" + vbCrLf) + _
                                   IIf((eResult And eavfPostcode), "", vbTab + "Postcode" + vbCrLf), _
                                vbInformation + vbOKOnly, _
                                "WARNING!"
                    End If

                Case etbStreet
                    .Street = txtDialog(etbStreet).Text
                    txtDialog(etbStreet).Text = .Street     '## reformats keyed field
                    txtDialog(etbAddress).Text = .Address   '## returns formatted name

                Case etbCity
                    .City = txtDialog(etbCity).Text
                    txtDialog(etbCity).Text = .City         '## reformats keyed field
                    txtDialog(etbAddress).Text = .Address   '## returns formatted name

                Case etbState
                    .State = txtDialog(etbState).Text
                    txtDialog(etbState).Text = .State       '## reformats keyed field
                    txtDialog(etbAddress).Text = .Address   '## returns formatted name

                Case etbPostcode
                    .Postcode = txtDialog(etbPostcode).Text
                    txtDialog(etbPostcode).Text = .Postcode '## reformats keyed field
                    txtDialog(etbAddress).Text = .Address   '## returns formatted name

            End Select
        End With
        mbIsBusy = False    '## re-enable the app's VB event handling
    End If
    mbIsDirty = False       '## Changes applied, therefore reset dirty flag

End Sub
