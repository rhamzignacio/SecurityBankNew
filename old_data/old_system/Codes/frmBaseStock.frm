VERSION 5.00
Begin VB.Form frmBaseStock 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Basestock"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6810
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBaseStock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.Frame Frame3 
         BackColor       =   &H00404000&
         Caption         =   "Paper Format"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   6375
         Begin VB.OptionButton optCutSheet 
            BackColor       =   &H00404000&
            Caption         =   "Cut Sheet"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   720
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton optContinues 
            BackColor       =   &H00404000&
            Caption         =   "Continues"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   2880
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00404000&
         Caption         =   "Paper Size"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   6375
         Begin VB.OptionButton opt8Outs 
            BackColor       =   &H00404000&
            Caption         =   "8 Outs"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   2880
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton opt4Outs 
            BackColor       =   &H00404000&
            Caption         =   "4 Outs"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   720
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Label lblJobName 
         BackColor       =   &H00404000&
         Caption         =   "Job Name:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404000&
         Caption         =   "Job Name:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmBaseStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Basestock_ChkType As String
Dim Basestock_FormType As String

Dim Basestock_ChkType_2 As String
Dim Basestock_FormType_2 As String



Sub LoadCheckType()
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT ChkType, FormType FROM SBTC WHERE BStock IS NULL  ORDER BY ChkType, FormType"
dbfRecordset.Open SQL, DBFConnector, 1, 1


If dbfRecordset.RecordCount >= 1 Then
    Basestock_ChkType = dbfRecordset.Fields(0)
    Basestock_FormType = dbfRecordset.Fields(1)
    
    Basestock_ChkType_2 = Basestock_ChkType
    Basestock_FormType_2 = Basestock_FormType
    
    
    
    If Basestock_ChkType = "A" And Basestock_FormType = "05" Then
        Basestock_ChkType_2 = "AA"
        Basestock_FormType_2 = "05"
    End If
    
    If Basestock_ChkType = "B" And Basestock_FormType = "16" Then
        Basestock_ChkType_2 = "BB"
        Basestock_FormType_2 = "16"
    End If


    
    
    If Basestock_ChkType = "A" And Basestock_FormType = "05" Then JobName = "Personal Checks"
    If Basestock_ChkType = "B" And Basestock_FormType = "16" Then JobName = "Commercial Checks"
    
    If Basestock_ChkType = "AA" And Basestock_FormType = "05" Then JobName = "Personal Checks"
    If Basestock_ChkType = "BB" And Basestock_FormType = "16" Then JobName = "Commercial Checks"
    
    If Basestock_ChkType = "MC" And Basestock_FormType = "20" Then JobName = "Manager's Checks"
    
    If Basestock_ChkType = "F" And Basestock_FormType = "25" Then JobName = "CheckOne Personal Checks"
    If Basestock_ChkType = "F" And Basestock_FormType = "26" Then JobName = "CheckOne Commercial Checks"
    
    If Basestock_ChkType = "E" And Basestock_FormType = "23" Then JobName = "CheckPower Personal Checks"
    If Basestock_ChkType = "E" And Basestock_FormType = "22" Then JobName = "CheckPower Commercial Checks"
    
    If Basestock_ChkType = "GC" And Basestock_FormType = "20" Then JobName = "Gift Check"
    
        
    
    lblJobName.Caption = JobName
Else
    Basetock_Verify = InputBox("Enter the name of the screeners", "", "")
    If Basetock_Verify = "" Then End
    
    
    Unload Me
    frmMain.Show
    
    
    frmMain.cmdCheckFilesHead.Caption = "Process ! ! !"
End If

End Sub


Private Sub cmdOk_Click()
If opt4Outs.Value = False And opt8Outs.Value = False Then
    MsgBox "Please select paper size", vbInformation, "Error"
    Exit Sub
End If



If optCutsheet.Value = False And optContinues.Value = False Then
    MsgBox "Please select paper format", vbInformation, "Error"
    Exit Sub
End If




If opt4Outs.Value = True Then PaperSize = "4 Outs"
If opt8Outs.Value = True Then PaperSize = "8 Outs"




If optCutsheet.Value = True Then PaperFormat = "Cut Sheet"
If optContinues.Value = True Then PaperFormat = "Continues"




BStock = PaperSize & " (" & PaperFormat & ")"




Set DBFConnector = CreateObject("ADODB.Connection")
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient



Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "UPDATE SBTC SET BStock = '" & BStock & "' WHERE (ChkType = '" & Basestock_ChkType & "' AND FormType = '" & Basestock_FormType & "') OR (ChkType = '" & Basestock_ChkType_2 & "' AND FormType = '" & Basestock_FormType_2 & "')"
dbfRecordset.Open SQL, DBFConnector, 1, 1




Unload Me
Me.Show


LoadCheckType


End Sub

