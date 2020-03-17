VERSION 5.00
Begin VB.Form frmEncode 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encode Orders"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11100
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEncode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   11100
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00404000&
      Caption         =   "Pcs Per Book"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   5520
      Width           =   10815
      Begin VB.OptionButton opt50 
         BackColor       =   &H00404000&
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton opt100 
         BackColor       =   &H00404000&
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   2280
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame grpPly 
      BackColor       =   &H00404000&
      Caption         =   "Duplicate Copy"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   4800
      TabIndex        =   24
      Top             =   4680
      Width           =   6135
      Begin VB.OptionButton opt4Ply 
         BackColor       =   &H00404000&
         Caption         =   "4 Ply"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   4440
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton opt3Ply 
         BackColor       =   &H00404000&
         Caption         =   "3 Ply"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   2880
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton opt2Ply 
         BackColor       =   &H00404000&
         Caption         =   "2 Ply"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   1560
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton opt1Ply 
         BackColor       =   &H00404000&
         Caption         =   "1 Ply"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame grpBasestock 
      BackColor       =   &H00404000&
      Caption         =   "Basestock"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   4680
      Width           =   4575
      Begin VB.OptionButton optContinues 
         BackColor       =   &H00404000&
         Caption         =   "Continues"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   2280
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optCutsheet 
         BackColor       =   &H00404000&
         Caption         =   "Cut Sheet"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Process"
      Height          =   495
      Left            =   9360
      TabIndex        =   9
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   7440
      Width           =   1575
   End
   Begin VB.ListBox lstDisplay 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      TabIndex        =   18
      Top             =   6360
      Width           =   10815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Height          =   4455
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   10815
      Begin VB.TextBox txtStartingSerial 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   7
         Top             =   3840
         Width           =   3135
      End
      Begin VB.TextBox txtAccountNo 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   3
         Top             =   2040
         Width           =   8415
      End
      Begin VB.TextBox txtOrderQty 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   6
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtName2 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   5
         Top             =   3240
         Width           =   8415
      End
      Begin VB.TextBox txtName1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   4
         Top             =   2640
         Width           =   8415
      End
      Begin VB.TextBox txtBRSTN 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         MaxLength       =   9
         TabIndex        =   2
         Top             =   1440
         Width           =   8415
      End
      Begin VB.ComboBox cboBranchName 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   8415
      End
      Begin VB.ComboBox cboChkType 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   8415
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404000&
         Caption         =   "Starting Serial:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   5520
         TabIndex        =   20
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404000&
         Caption         =   "Account No:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404000&
         Caption         =   "100 Pcs / Bkt"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   3600
         TabIndex        =   17
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404000&
         Caption         =   "Books:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404000&
         Caption         =   "Account Name 2:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404000&
         Caption         =   "Account Name 1:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404000&
         Caption         =   "BRSTN:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblBranchName 
         BackColor       =   &H00404000&
         Caption         =   "Branch Name:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404000&
         Caption         =   "Cheque Type:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Label lblDeliveryDate 
      BackColor       =   &H00404000&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   4200
      TabIndex        =   32
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H00404000&
      Caption         =   "Delivery Date:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   2280
      TabIndex        =   33
      Top             =   7440
      Width           =   2055
   End
End
Attribute VB_Name = "frmEncode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboBranchName_Click()
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC\Continues\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT BRSTN FROM Branches WHERE Address1 = '" & Replace(cboBranchName.Text, "'", "''") & "'"
dbfRecordset.Open SQL, DBFConnector, 1, 1

txtBRSTN.Text = dbfRecordset.Fields(0)
End Sub

Private Sub cboChkType_Click()
optCutsheet.Value = False
optContinues.Value = False

opt1Ply.Value = False
opt2Ply.Value = False
opt3Ply.Value = False
opt4Ply.Value = False

opt50.Value = False
opt100.Value = False




If cboChkType.Text = "MC Continues" Then
    cboBranchName.Enabled = True
    txtBRSTN.Text = ""
    txtBRSTN.Enabled = False
    
    txtAccountNo.Text = ""
    txtAccountNo.Enabled = True
    
    LoadBranchName
    
    
    grpBasestock.Enabled = False
    grpPly.Enabled = False
    
    'Pcs Per Book
    Frame2.Enabled = False
    opt100.Value = True
    'End Pcs Per Book
Else
    If cboChkType.Text = "Charge Slip" Then
        cboBranchName.Clear
        cboBranchName.Enabled = False
        
        txtBRSTN.Text = "010140455"
        txtBRSTN.Enabled = False
        
        txtAccountNo.Text = "000000000000"
        txtAccountNo.Enabled = False
        
        
        'Pcs Per Book
        Frame2.Enabled = False
        opt50.Value = True
        
        optCutsheet.Value = True
        opt2Ply.Value = True
        'End Pcs Per Book
    Else
        txtAccountNo.Text = ""
        txtAccountNo.Enabled = True
    
        cboBranchName.Clear
        cboBranchName.Enabled = False
        txtBRSTN.Enabled = True
        
        
        grpBasestock.Enabled = True
        grpPly.Enabled = True
        
        
        'Pcs Per Book
        Frame2.Enabled = True
        opt50.Value = False
        opt100.Value = False
        'End Pcs Per Book
    End If

End If
End Sub

Sub LoadBranchName()
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC\Continues\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Address1 FROM Branches ORDER BY Address1"
dbfRecordset.Open SQL, DBFConnector, 1, 1

cboBranchName.Clear

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    Address1 = dbfRecordset.Fields(0)
        
    cboBranchName.AddItem (Address1)
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
End Sub

Private Sub cmdAdd_Click()



Set DBFConnector = CreateObject("ADODB.Connection")
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient



If cboChkType.Text = "" Then
    MsgBox "Cheque Type is Invalid", vbInformation, "Error"
    cboChkType.SetFocus
    Exit Sub
End If

If cboBranchName.Text = "" And cboBranchName.Enabled = True Then
    MsgBox "Branch Name is Invalid", vbInformation, "Error"
    cboBranchName.SetFocus
    Exit Sub
End If

If (Len(txtBRSTN.Text) <> 9 Or IsNumeric(txtBRSTN.Text) = False) And txtBRSTN.Enabled = True Then
    MsgBox "BRSTN is Invalid", vbInformation, "Error"
    txtBRSTN.SetFocus
    Exit Sub
End If

If Len(txtAccountNo.Text) <> 12 Or IsNumeric(txtAccountNo.Text) = False Then
    MsgBox "Account Number is Invalid", vbInformation, "Error"
    txtAccountNo.SetFocus
    Exit Sub
End If

If Val(txtOrderQty.Text) = 0 Then
    MsgBox "Order Qty is Invalid", vbInformation, "Error"
    txtOrderQty.SetFocus
    Exit Sub
End If

If Val(txtStartingSerial.Text) = 0 Then
    MsgBox "Starting Serial is Invalid", vbInformation, "Error"
    txtStartingSerial.SetFocus
    Exit Sub
End If




'For Pcs Per Book
If opt50.Value = False And opt100.Value = False Then
    MsgBox "Please select Pcs per Book", vbInformation, "Error"
    Exit Sub
End If

If opt50.Value = True Then PcsPerBook = 50
If opt100.Value = True Then PcsPerBook = 100
'End For Pcs Per Book







If cboChkType.Text = "Customized" Then
    Set dbfRecordset = CreateObject("ADODB.Recordset")
    SQL = "SELECT Address1 , Address2 , Address3 , Address4 , Address5 , Address6 FROM Branches WHERE BRSTN = '" & txtBRSTN.Text & "'"
    dbfRecordset.Open SQL, DBFConnector, 1, 1
    
    
    If dbfRecordset.RecordCount = 0 Then
        MsgBox "BRSTN " & txtBRSTN.Text & " does not exists on Branches.dbf", vbInformation, "Error"
        txtBRSTN.Text = ""
        txtBRSTN.SetFocus
        Exit Sub
    End If
    
    If Len(dbfRecordset.Fields(0)) >= 1 Then
        Address1 = dbfRecordset.Fields(0)
    Else
        Address1 = ""
    End If

    If Len(dbfRecordset.Fields(1)) >= 1 Then
        Address2 = dbfRecordset.Fields(1)
    Else
        Address2 = ""
    End If
    
    If Len(dbfRecordset.Fields(2)) >= 1 Then
        Address3 = dbfRecordset.Fields(2)
    Else
        Address3 = ""
    End If
    
    If Len(dbfRecordset.Fields(3)) >= 1 Then
        Address4 = dbfRecordset.Fields(3)
    Else
        Address4 = ""
    End If
    
    If Len(dbfRecordset.Fields(4)) >= 1 Then
        Address5 = dbfRecordset.Fields(4)
    Else
        Address5 = ""
    End If
    
    If Len(dbfRecordset.Fields(5)) >= 1 Then
        Address6 = dbfRecordset.Fields(5)
    Else
        Address6 = ""
    End If
    
    
    
    
    If optCutsheet.Value = False And optContinues.Value = False Then
        MsgBox "Please select Basestock", vbInformation, "Error"
        Exit Sub
    End If
    
    If opt1Ply.Value = False And opt2Ply.Value = False And opt3Ply.Value = False And opt4Ply.Value = False Then
        MsgBox "Please select Duplicate copy", vbInformation, "Error"
        Exit Sub
    End If
            
    
    'For Basestock
    BStock = ""
    If optCutsheet.Value = True Then BStock = "Cutsheet"
    If optContinues.Value = True Then BStock = "Continues"
    'End For Basestock
    
    
    
    'Duplicate
    Duplicate = ""
    
    If opt1Ply.Value = True Then BStock = BStock & "(1 Ply)"
    If opt2Ply.Value = True Then BStock = BStock & "(2 Ply)"
    If opt3Ply.Value = True Then BStock = BStock & "(3 Ply)"
    If opt4Ply.Value = True Then BStock = BStock & "(4 Ply)"
    'End Duplicate
    
    ChkType = "CUSTOM"
    FormType = "00"
End If



If cboChkType.Text = "MC Continues" Then
    ChkType = "MC_1"
    FormType = "00"
    BStock = "Continues"
    
    
    Set DBFConnector1 = CreateObject("ADODB.Connection")
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC\Continues\;Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient

    'For Address
    Set dbfRecordset = CreateObject("ADODB.Recordset")
    SQL = "SELECT Address2, Address3, Address4, Address5, Address6 FROM Branches WHERE Address1 = '" & Replace(cboBranchName.Text, "'", "''") & "'"
    dbfRecordset.Open SQL, DBFConnector1, 1, 1

    Address1 = cboBranchName.Text
    
    If Len(dbfRecordset.Fields(0)) >= 1 Then
        Address2 = dbfRecordset.Fields(0)
    Else
        Address2 = ""
    End If
    
    If Len(dbfRecordset.Fields(1)) >= 1 Then
        Address3 = dbfRecordset.Fields(1)
    Else
        Address3 = ""
    End If
    
    If Len(dbfRecordset.Fields(2)) >= 1 Then
        Address4 = dbfRecordset.Fields(2)
    Else
        Address4 = ""
    End If
    
    If Len(dbfRecordset.Fields(3)) >= 1 Then
        Address5 = dbfRecordset.Fields(3)
    Else
        Address5 = ""
    End If
    
    If Len(dbfRecordset.Fields(4)) >= 1 Then
        Address6 = dbfRecordset.Fields(4)
    Else
        Address6 = ""
    End If
    'End For Address
End If





If cboChkType.Text = "Charge Slip" Then
    ChkType = "CS"
    Address1 = "CHARGE SLIP"
    FormType = "00"
End If






Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "INSERT INTO SBTC (ChkType, BRSTN, AccountNo, Name1, Name2, OrderQty,FormType, Address1, Address2, Address3, Address4, Address5, Address6, BStock, StartSN , PcsPerBook , pkey) VALUES ('" _
    & ChkType & "','" & txtBRSTN.Text & "','" & txtAccountNo.Text & "','" & UCase(Replace(txtName1.Text, "'", "''")) & "','" & UCase(Replace(txtName2.Text, "'", "''")) & "','" & txtOrderQty.Text & "','" & FormType & "','" & Replace(Address1, "'", "''") & "','" & Replace(Address2, "'", "''") & "','" & Replace(Address3, "'", "''") & "','" & Replace(Address4, "'", "''") & "','" & Replace(Address5, "'", "''") & "','" & Replace(Address6, "'", "''") & "','" & BStock & "','" & txtStartingSerial.Text & "','" & PcsPerBook & "'," & getPKey & ")"
dbfRecordset.Open SQL, DBFConnector, 1, 1







If cboChkType.Text = "MC Continues" Then SortRT ("MC\Continues")
If cboChkType.Text = "Customized" Then SortRT ("Customized")
If cboChkType.Text = "Charge Slip" Then SortRT ("Charge_Slip")






MsgBox "Data has been Saved", vbInformation, ""

DeliveryDate = lblDeliveryDate.Caption

Unload Me
Me.Show

lblDeliveryDate.Caption = DeliveryDate
End Sub

Private Sub cmdProcess_Click()
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient


Batch = InputBox("Enter Batch Name", "", Format(Now, "MMDDYYYY"))
If Batch = "" Then Exit Sub

If Dir$(App.Path & "\Archive\" & Batch, vbDirectory) <> "" Then
    MsgBox "Batch " & Batch & " has already been processed", vbInformation, "Error"
    Exit Sub
End If


Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "UPDATE SBTC Set Batch = '" & Replace(Batch, "'", "''") & "'"
dbfRecordset.Open SQL, DBFConnector, 1, 1




'Process By / Checked By
ProcessBy = InputBox("Enter Process By", "", "")
If ProcessBy = "" Then Exit Sub

CheckedBy = InputBox("Enter Checked By", "", "")
If CheckedBy = "" Then Exit Sub

If UCase(ProcessBy) = UCase(CheckedBy) Then
    MsgBox "Prepared by and Checked By should not be the same", vbCritical, "Error"
    Exit Sub
End If
'End Process By / Checked By



DateTimeToday = Format(Now, "HHMMSSMMDDYY")
MkDir ("C:\Windows\Temp\" & DateTimeToday)
MkDir ("C:\Windows\Temp\" & DateTimeToday & "\" & Batch)

Result = ProcessAll(Batch, ProcessBy, CheckedBy, lblDeliveryDate.Caption)




MsgBox "Data has been processed", vbInformation, ""
End

    
End Sub

Private Sub Form_Load()
loadChkType

LoadDisplay
End Sub


Sub LoadDisplay()
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT ChkType, BRSTN, AccountNo, OrderQty, Name1, Name2 FROM SBTC"
dbfRecordset.Open SQL, DBFConnector, 1, 1

lstDisplay.Clear

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    ChkType = dbfRecordset.Fields(0)
    BRSTN = dbfRecordset.Fields(1)
    AccountNo = dbfRecordset.Fields(2)
    OrderQty = dbfRecordset.Fields(3)
    
    If Len(dbfRecordset.Fields(4)) >= 1 Then
        Name1 = dbfRecordset.Fields(4)
    Else
        Name1 = ""
    End If
    
    If Len(dbfRecordset.Fields(5)) >= 1 Then
        Name2 = dbfRecordset.Fields(5)
    Else
        Name2 = ""
    End If
    
    Do Until Len(ChkType) >= 10
        ChkType = ChkType & " "
    Loop
    
    Do Until Len(OrderQty) >= 6
        OrderQty = OrderQty & " "
    Loop
    
    lstDisplay.AddItem (ChkType & BRSTN & "   " & AccountNo & "   " & OrderQty & Name1 & " " & Name2)
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

If dbfRecordset.RecordCount >= 1 Then
    cmdProcess.Enabled = True
Else
    cmdProcess.Enabled = False
End If
End Sub


Sub loadChkType()
cboChkType.Clear
cboChkType.AddItem ("Customized")
cboChkType.AddItem ("MC Continues")
cboChkType.AddItem ("Charge Slip")
End Sub
