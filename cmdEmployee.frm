VERSION 5.00
Begin VB.Form cmdEmployee 
   BackColor       =   &H80000000&
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtdatehired 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   16
      Top             =   3000
      Width           =   3735
   End
   Begin VB.TextBox txtsalary 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   14
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox txtposition 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   12
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox txtaddress 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   3735
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox txtEmployeeID 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton closebttn 
      Caption         =   "CLOSE"
      Height          =   975
      Left            =   5400
      Picture         =   "cmdEmployee.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton findbttn 
      Caption         =   "FIND"
      Height          =   975
      Left            =   4080
      Picture         =   "cmdEmployee.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton deletebttn 
      Caption         =   "DELETE"
      Height          =   975
      Left            =   2880
      Picture         =   "cmdEmployee.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton savebttn 
      Caption         =   "SAVE"
      Height          =   975
      Left            =   1560
      Picture         =   "cmdEmployee.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton addbttn 
      BackColor       =   &H8000000B&
      Caption         =   "ADD"
      Height          =   975
      Left            =   240
      Picture         =   "cmdEmployee.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000000&
      Caption         =   "DATE HIRED:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000000&
      Caption         =   "SALARY:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   13
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000000&
      Caption         =   "POSITION:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000000&
      Caption         =   "ADDRESS:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label dfsz 
      BackColor       =   &H80000000&
      Caption         =   "NAME:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label txtid 
      BackColor       =   &H80000000&
      Caption         =   "EMPLOYEE ID:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "cmdEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

End Sub

Private Sub addbttn_Click()
txtEmployeeID.Text = ""
txtname.Text = ""
txtaddress.Text = ""
txtposition.Text = ""
txtsalary.Text = ""
txtdatehired.Text = ""
txtEmployeeID.SetFocus
End Sub

Private Sub closebttn_Click()
 Unload Me
End Sub

Private Sub deletebttn_Click()
conPayroll.Execute "Delete * from employee where employeeid='" & Trim(txtEmployeeID.Text) & "'"
MsgBox "Record has been deleted.."
End Sub

Private Sub findbttn_Click()
txtEmployeeID.SelStart = 0
txtEmployeeID.SelLength = Len(txtEmployeeID.Text)
txtEmployeeID.SetFocus
End Sub

Private Sub Form_Load()
openWORKSPACEODBC
openconPayroll
End Sub

Private Sub savebttn_Click()
openrstEmployee "Select * from EMPLOYEE where employeeid='" & Trim(txtEmployeeID.Text) & "'"
If Not rstEmployee.EOF Then
    'not found
    With rstEmployee
        .Edit
            .Fields("employeeid").Value = txtEmployeeID.Text
            .Fields("employeename").Value = txtname.Text
            .Fields("address").Value = txtaddress.Text
            .Fields("position").Value = txtposition.Text
            .Fields("salary").Value = txtsalary.Text
            .Fields("datehired").Value = txtdatehired.Text
            
        .Update
        
        
    End With
Else
    'found
        With rstEmployee
        .AddNew
             .Fields("employeeid").Value = txtEmployeeID.Text
            .Fields("employeename").Value = txtname.Text
            .Fields("address").Value = txtaddress.Text
            .Fields("position").Value = txtposition.Text
            .Fields("salary").Value = txtsalary.Text
            .Fields("datehired").Value = txtdatehired.Text
        .Update
        
        End With
End If
End Sub

Private Sub txtaddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtposition.SetFocus
    End If
End Sub

Private Sub txtEmployeeID_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    openrstEmployee "Select * from EMPLOYEE where employeeid ='" & Trim(txtEmployeeID.Text) & "'"
     If Not rstEmployee.EOF Then
        With rstEmployee
            txtEmployeeID.Text = .Fields("employeeid").Value
            txtname.Text = .Fields("employeename").Value
            txtaddress.Text = .Fields("address").Value
            txtposition.Text = .Fields("position").Value
            txtsalary.Text = .Fields("salary").Value
            txtdatehired.Text = .Fields("datehired").Value
        End With
    End If
    
End If
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtaddress.SetFocus
    End If
End Sub

Private Sub txtposition_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtsalary.SetFocus
    End If
End Sub

Private Sub txtsalary_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtdatehired.SetFocus
    End If
End Sub
