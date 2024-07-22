VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmmdi 
   BackColor       =   &H80000006&
   Caption         =   "MDIForm1"
   ClientHeight    =   3570
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8520
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "Employee"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":050E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   1588
      ButtonWidth     =   1561
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Employee"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Payroll"
            ImageIndex      =   3
         EndProperty
      EndProperty
      MouseIcon       =   "frmMain.frx":0960
   End
   Begin VB.Menu menufile 
      Caption         =   "&File"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu menuemployee 
         Caption         =   "&Employee"
      End
      Begin VB.Menu menuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu menutrans 
      Caption         =   "&Transaction"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu menupayroll 
         Caption         =   "&Payroll"
      End
   End
   Begin VB.Menu menureps 
      Caption         =   "&Reports"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu menusummaryearn 
         Caption         =   "&SummaryEarnings"
      End
   End
End
Attribute VB_Name = "frmmdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub menuemployee_Click()
cmdEmployee.Show
End Sub

Private Sub menuexit_Click()
Unload Me
End Sub

Private Sub menupayroll_Click()
cmdPayroll.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
        Case "1"
        cmdEmployee.Show
        Case "2"
        cmdPayroll.Show
        
        End Select
        
End Sub
