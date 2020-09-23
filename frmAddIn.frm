VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Code Line Counter"
   ClientHeight    =   3480
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6600
   Icon            =   "frmAddIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstReport 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   1500
      TabIndex        =   14
      ToolTipText     =   "Analysis of the code in the selected component"
      Top             =   2130
      Width           =   2685
   End
   Begin VB.ComboBox cboProject 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Click to select an open project"
      Top             =   330
      Width           =   2625
   End
   Begin VB.ComboBox cboComponent 
      Height          =   315
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Click to select a component in the project"
      Top             =   1710
      Width           =   2625
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   525
      Left            =   5550
      TabIndex        =   2
      ToolTipText     =   "Click when you are finished"
      Top             =   390
      Width           =   1005
   End
   Begin VB.Label lblDesigners 
      Caption         =   "0"
      Height          =   225
      Left            =   4260
      TabIndex        =   18
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Designers:"
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   945
      Width           =   1395
   End
   Begin VB.Label Label6 
      Caption         =   "User Controls:"
      Height          =   255
      Left            =   2640
      TabIndex        =   16
      Top             =   705
      Width           =   1395
   End
   Begin VB.Label lblUserControls 
      Caption         =   "0"
      Height          =   225
      Left            =   4260
      TabIndex        =   15
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblComponentLines 
      Height          =   225
      Left            =   4500
      TabIndex        =   13
      Top             =   1770
      Width           =   705
   End
   Begin VB.Label lblProjectLines 
      Height          =   225
      Left            =   4500
      TabIndex        =   12
      Top             =   360
      Width           =   705
   End
   Begin VB.Label lblClassModules 
      Caption         =   "0"
      Height          =   225
      Left            =   1980
      TabIndex        =   11
      Top             =   990
      Width           =   495
   End
   Begin VB.Label lblForms 
      Caption         =   "0"
      Height          =   225
      Left            =   1980
      TabIndex        =   10
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblCodeModules 
      Caption         =   "0"
      Height          =   225
      Left            =   1980
      TabIndex        =   9
      Top             =   750
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Class Modules:"
      Height          =   225
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label Label3 
      Caption         =   "Forms:"
      Height          =   225
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "Code Modules:"
      Height          =   225
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "Code Lines"
      Height          =   225
      Left            =   4410
      TabIndex        =   5
      Top             =   90
      Width           =   885
   End
   Begin VB.Label lblComponent 
      Caption         =   "Component"
      Height          =   225
      Left            =   60
      TabIndex        =   1
      Top             =   1770
      Width           =   1215
   End
   Begin VB.Label lblProject 
      Caption         =   "Project"
      Height          =   225
      Left            =   90
      TabIndex        =   0
      Top             =   360
      Width           =   1275
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Option Explicit

Private lngDeclarationLines As Long
Private lngBlankLines As Long
Private lngCodeLines As Long
Private lngCommentLines As Long
Private lngContinuedLines As Long

Private Sub cboComponent_Click()
  Dim vbpProject As VBProject
  
  Set vbpProject = VBInstance.VBProjects.Item(Me.cboProject)
  Me.lblComponentLines = vbpProject.VBComponents.Item(Me.cboComponent).CodeModule.CountOfLines
  
  parseModule vbpProject.VBComponents.Item(Me.cboComponent).CodeModule
  
  Me.lstReport.Clear
  Me.lstReport.AddItem "Declaration Lines:= " & lngDeclarationLines
  Me.lstReport.AddItem "Blank Lines:= " & lngBlankLines
  Me.lstReport.AddItem "Code Lines:= " & lngCodeLines
  Me.lstReport.AddItem "Comment Lines:= " & lngCommentLines
  Me.lstReport.AddItem "Continuation Lines:= " & lngContinuedLines

End Sub

Private Sub cboProject_Click()
  Dim comComponent As VBProject
  Dim vbpProject As VBProject
  Dim intLoop As Integer
  Dim lngTotalLines As Long
  Dim intModuleCount As Integer, intClassCount As Integer, intFormCount As Integer, intUserControlCount As Integer, intDesignerCount As Integer
  lngTotalLines = 0
  intModuleCount = 0
  intClassCount = 0
  intFormCount = 0
  intUserControlCount = 0
  intDesignerCount = 0
  Me.lstReport.Clear
  
  Set vbpProject = VBInstance.VBProjects.Item(Me.cboProject)
  cboComponent.Clear
  For intLoop = 1 To vbpProject.VBComponents.Count
    cboComponent.AddItem vbpProject.VBComponents(intLoop).Name
    lngTotalLines = lngTotalLines + vbpProject.VBComponents(intLoop).CodeModule.CountOfLines
    Select Case vbpProject.VBComponents(intLoop).Type
      Case vbext_ct_ClassModule
        intClassCount = intClassCount + 1
      Case vbext_ct_MSForm
        intFormCount = intFormCount + 1
      Case vbext_ct_StdModule
        intModuleCount = intModuleCount + 1
      Case vbext_ct_UserControl
        intUserControlCount = intUserControlCount + 1
      Case vbext_ct_VBForm
        intFormCount = intFormCount + 1
      Case vbext_ct_VBMDIForm
        intFormCount = intFormCount + 1
      Case vbext_ct_ActiveXDesigner
        intDesignerCount = intDesignerCount + 1
    End Select
  Next intLoop
  Me.lblProjectLines = lngTotalLines
  Me.lblCodeModules = intModuleCount
  Me.lblClassModules = intClassCount
  Me.lblForms = intFormCount
  Me.lblUserControls = intUserControlCount
  Me.lblDesigners = intDesignerCount
End Sub

Private Sub cmdDone_Click()

'Hide the AddIn window
Connect.Hide

End Sub

Private Sub Form_Load()
  Dim vbpProject As VBProject
  Dim intLoop As Integer
  cboProject.Clear
  cboComponent.Clear
  For intLoop = 1 To VBInstance.VBProjects.Count
    cboProject.AddItem VBInstance.VBProjects(intLoop).Name
  Next intLoop
End Sub

Private Function parseModule(codModule As CodeModule)
   Dim lngLines As Long, lngLine As Long
   Dim lngCharacter As Long
   Dim strLine As String
   
   lngDeclarationLines = codModule.CountOfDeclarationLines
   lngBlankLines = 0
   lngCodeLines = 0
   lngCommentLines = 0
   lngContinuedLines = 0
   
   lngLines = codModule.CountOfLines
   For lngLine = 1 To lngLines
     strLine = codModule.Lines(lngLine, 1)
     If Len(strLine) = 0 Then
       'blank line
       lngBlankLines = lngBlankLines + 1
       'remove one from declarations count if in declarations area
       If lngLine <= codModule.CountOfDeclarationLines Then lngDeclarationLines = lngDeclarationLines - 1
     Else
       If Len(Trim(strLine)) = 0 Then
         'blank line
         lngBlankLines = lngBlankLines + 1
         If lngLine <= codModule.CountOfDeclarationLines Then lngDeclarationLines = lngDeclarationLines - 1
       Else
         'comment line
         If Left(LTrim(strLine), 1) = "'" Then
           lngCommentLines = lngCommentLines + 1
           If lngLine <= codModule.CountOfDeclarationLines Then lngDeclarationLines = lngDeclarationLines - 1
         Else
           'continuation line
           If lngLine > 1 Then
             If Right(RTrim(codModule.Lines(lngLine - 1, 1)), 1) = "_" Then
               lngContinuedLines = lngContinuedLines + 1
               If lngLine <= codModule.CountOfDeclarationLines Then lngDeclarationLines = lngDeclarationLines - 1
             Else
               'code line
               If lngLine > codModule.CountOfDeclarationLines Or codModule.CountOfDeclarationLines = 0 Then
                 lngCodeLines = lngCodeLines + 1
               End If
             End If
           Else
             If lngLine > codModule.CountOfDeclarationLines Or codModule.CountOfDeclarationLines = 0 Then
               lngCodeLines = lngCodeLines + 1
             End If
           End If
         End If
       End If
     End If
   Next lngLine
End Function

