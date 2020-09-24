VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "rion Belt - Proceeding"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProcess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar prgOverall 
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   9
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar prgCurrent 
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   735
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   9
      Scrolling       =   1
   End
   Begin VB.Label cmdNo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "&NO!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   2415
      TabIndex        =   12
      Top             =   3495
      Width           =   225
   End
   Begin VB.Shape shpButtonBorder 
      BorderColor     =   &H00808080&
      Height          =   165
      Index           =   1
      Left            =   2400
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label cmdYes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "&YES!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1230
      TabIndex        =   11
      Top             =   3375
      Width           =   1080
   End
   Begin VB.Shape shpButtonBorder 
      BorderColor     =   &H00808080&
      Height          =   345
      Index           =   0
      Left            =   1200
      Top             =   3360
      Width           =   1125
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Are You Sure You Want To Do This?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label lblConfirm 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmProcess.frx":324A
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label lblCurrent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Idle - Waiting for comfirmation..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label lblCurrent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Idle - Waiting for comfirmation..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   30
      TabIndex        =   8
      Top             =   1350
      Width           =   4095
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   272
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   272
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Label lblPercentage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Operation: 0%"
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   975
      Width           =   1695
   End
   Begin VB.Label lblPercentage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Overall Operation Completed: 0%"
      Height          =   195
      Index           =   0
      Left            =   810
      TabIndex        =   2
      Top             =   360
      Width           =   2475
   End
   Begin VB.Shape shpPogressBorder 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   1
      Left            =   105
      Top             =   720
      Width           =   3840
   End
   Begin VB.Shape shpPogressBorder 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   0
      Left            =   105
      Top             =   105
      Width           =   3840
   End
   Begin VB.Shape shpProgressShadow 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   0
      Left            =   135
      Top             =   195
      Width           =   3855
   End
   Begin VB.Shape shpProgressShadow 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   1
      Left            =   135
      Top             =   810
      Width           =   3855
   End
   Begin VB.Label lblShadow 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Overall Operation Completed: 0%"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   0
      Left            =   825
      TabIndex        =   4
      Top             =   375
      Width           =   2475
   End
   Begin VB.Label lblShadow 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Operation: 0%"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   1215
      TabIndex        =   5
      Top             =   990
      Width           =   1695
   End
   Begin VB.Label lblCurrentBackground 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Shape shpButtonShadow 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   315
      Index           =   0
      Left            =   1305
      Top             =   3450
      Width           =   1080
   End
   Begin VB.Shape shpButtonShadow 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   2460
      Top             =   3540
      Width           =   225
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'   ============================================================
'    ----------------------------------------------------------
'     Application Name: Orion Belt
'                       The World's Best VB Code Commenter
'     Developer/Programmer: Alph@
'    ----------------------------------------------------------
'     Module Name: frmProgress
'     Module File: frmProcess.frm
'     Module Type: Form
'     Module Description:
'    ----------------------------------------------------------
'     ©opyright 2001-2002 by Alph@ - All Right Reserved
'    ----------------------------------------------------------
'   ============================================================

Option Explicit
Dim modpModules() As ModuleProperties
Dim bytModuleSize As Byte



'----------------------------------------
'Name: cmdNo_MouseDown
'Object: cmdNo
'Event: MouseDown
'----------------------------------------
Private Sub cmdNo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdNo.Left = 164
    cmdNo.Top = 236
    shpButtonBorder(1).Left = 163
    shpButtonBorder(1).Top = 235
    cmdNo.ZOrder
    shpButtonBorder(1).ZOrder
End Sub


'----------------------------------------
'Name: cmdNo_MouseUp
'Object: cmdNo
'Event: MouseUp
'----------------------------------------
Private Sub cmdNo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMain.Enabled = True
    Unload Me
End Sub


'----------------------------------------
'Name: cmdYes_MouseDown
'Object: cmdYes
'Event: MouseDown
'----------------------------------------
Private Sub cmdYes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdYes.Left = 87
    cmdYes.Top = 350
    shpButtonBorder(0).Left = 86
    shpButtonBorder(0).Top = 350
    cmdYes.ZOrder
    shpButtonBorder(0).ZOrder
End Sub


'----------------------------------------
'Name: cmdYes_MouseUp
'Object: cmdYes
'Event: MouseUp
'----------------------------------------
Private Sub cmdYes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdYes.Left = 82
    cmdYes.Top = 345
    shpButtonBorder(0).Left = 81
    shpButtonBorder(0).Top = 344
    cmdYes.ZOrder
    shpButtonBorder(0).ZOrder
    DoEvents
    CommentMain
End Sub


'----------------------------------------
'Name: CommentMain
'----------------------------------------
Private Sub CommentMain()
    'Begin Commenting...
    '-------------------
    '  Ok, let's make it clear. There's 3 step in commenting the whole project.
    'The first step is to scan the .vbp project file for the list of all modules
    'in the project, and store it in an arrey. Next step, we'll scroll through
    'the list of the modules, and determine its type. Modules will also be
    'commented in this step. And then step 3, procedures will be comment,
    'according to the options user have selected. The program will repeat step 2
    'and 3 While Not all modules, according to the project file, is fully commented.
    '  This procedure will be the main loop, which call other step-procedures and
    'manage variables. Actually it can be put in the cmdYes_MouseUp sub but I
    'would prefer putting it in another sub like this, so it'd be easier to
    'change/call in newer versions.
    
    'Pre-Step: Set & Clear variables
    bytModuleSize = 1
    
    'Step 1 -----------------
    ReadProjectFile
    
    'Step 2 -----------------
    ScanModInfo
    
    'Step 3 -----------------
    CommentProject
    
    'Last Step: Notify User
    prgOverall.Value = 100
    prgCurrent.Value = 100
    DoEvents
    If MsgBox("Your project (" & IIf(modpModules(1).ModName = "Project1", frmMain.txtProjInfo(1), modpModules(1).ModName) & ") was fully commented!" & _
    IIf(frmMain.chkBackup.Value = 1, vbNewLine & "The backup project is stored in " & _
    Left$(strTargetProject, LastStr(1, strTargetProject, "\")) & "Backup", "") & _
    vbNewLine & "Do you want to terminate Orion Belt?", vbYesNo + vbQuestion, "Completed!") _
    = vbYes Then End 'Dunno, I've tried to unload every forms manually but it stills hang out in my (computer's) memory
    Unload Me
    frmMain.Enabled = True
End Sub


'----------------------------------------
'Name: ReadProjectFile
'----------------------------------------
Private Sub ReadProjectFile()
    Dim strInput As String
    Dim strKeyword As String
    Dim bytSeperator As Byte
    Dim booGotNothing As Boolean
    On Error Resume Next
    
    Open strTargetProject For Input As #1
        
        bytModuleSize = 1
        Do While Not EOF(1)
            Line Input #1, strInput
            If strInput <> "" Then If InStr(1, strInput, "=") = 0 Then strKeyword = "" Else strKeyword = Left$(strInput, InStr(1, strInput, "=") - 1)
            
            bytModuleSize = bytModuleSize + 1
            ReDim Preserve modpModules(1 To bytModuleSize) As ModuleProperties
            booGotNothing = False
            
            With modpModules(bytModuleSize)
                Select Case strKeyword
                Case "Name" 'Project
                    With modpModules(1)
                        .FileName = strTargetProject
                        .ModName = Mid$(strInput, Len(strKeyword) + 2)
                        .ModType = MProject
                    End With
                Case "Form" 'Form
                    .FileName = Mid$(strInput, Len(strKeyword) + 2)
                    .ModType = MForm
                Case "Module" 'Module
                    bytSeperator = InStr(Len(strKeyword) + 2, strInput, ";")
                    .FileName = Mid$(strInput, bytSeperator + 2)
                    .ModName = Mid$(strInput, Len(strKeyword) + 2, bytSeperator - Len(strKeyword) - 1)
                    .ModType = MModule
                Case "Class" 'Class Module
                    bytSeperator = InStr(Len(strKeyword) + 2, strInput, ";")
                    .FileName = Mid$(strInput, bytSeperator + 2)
                    .ModName = Mid$(strInput, Len(strKeyword) + 2, bytSeperator - Len(strKeyword) - 1)
                    .ModType = MClass
                Case "UserControl" 'ActiveX Control
                    .FileName = Mid$(strInput, Len(strKeyword) + 2)
                    .ModType = MUserControl
                Case "PropertyPage" 'Property Page
                    .FileName = Mid$(strInput, Len(strKeyword) + 2)
                    .ModType = MPropPage
                Case Else
                    booGotNothing = True
                End Select
            End With
            If booGotNothing Then
                bytModuleSize = bytModuleSize - 1
                ReDim Preserve modpModules(1 To bytModuleSize) As ModuleProperties
            End If
        Loop
        bytModuleSize = bytModuleSize - 1
    Close
End Sub


'----------------------------------------
'Name: ScanModInfo
'----------------------------------------
Private Sub ScanModInfo()
    Dim bytCount As Byte
    Dim strCurLine As String
    Dim booLoop As Boolean
    
    For bytCount = 2 To bytModuleSize
        Do
            If Dir$(modpModules(bytCount).FileName) <> "" Then
                Open modpModules(bytCount).FileName For Input As #1
                    Do While Not EOF(1)
                        Line Input #1, strCurLine
                        If Left$(strCurLine, 20) = "Attribute VB_Name = " Then
                            modpModules(bytCount).ModName = Left(Mid$(strCurLine, 22), Len(Mid$(strCurLine, 22)) - 1)
                            Exit Do
                        End If
                    Loop
                Close
            Else
                Select Case MsgBox("Orion Error 404: Path Not Found" & vbNewLine & _
                    modpModules(bytCount).FileName & " cannot be found." & vbNewLine & vbNewLine & _
                    "Click Abort if you want to cancel the whole operation to re-check your project." & vbNewLine & _
                    "Click Retry if you want Orion Belt to check for the file's existance again." & vbNewLine & _
                    "Click Ignore if you want to skip this file and continue with next file." _
                    , vbAbortRetryIgnore + vbCritical, "Error!")
                Case vbAbort
                    Unload Me
                Case vbRetry
                    booLoop = True
                Case vbIgnore
                    booLoop = False
                    modpModules(bytCount).FileName = ""
                End Select
            End If
        Loop While booLoop
    Next bytCount
End Sub


'----------------------------------------
'Name: CommentProject
'----------------------------------------
Private Sub CommentProject()
    Dim booCommentOne As Boolean
    Dim bytChar As Byte
    Dim bytCheck As Byte
    Dim bytCount As Byte
    Dim bytSeperator As Byte
    Dim strCode() As String
    Dim strCurLine As String
    Dim strBackupPath As String
    Dim strMajorDecor As String
    Dim strMinorDecor As String
    Dim strModuleType As String
    Dim intLine As Integer
    Dim intLineToIns As Integer
    
    'Initialize Backup Folder
    If frmMain.chkBackup.Value = 1 Then
        strBackupPath = Left$(strTargetProject, LastStr(1, strTargetProject, "\")) & "Backup"
        If Dir$(strBackupPath) <> "" Then
            If MsgBox("The backup folder is already exist. All the existing file in the folder must be deleted before Orion Belt could continue. Proceed?", vbYesNo + vbQuestion, "Duplicate Folder Name") Then
                RmDir strBackupPath
                MkDir strBackupPath
            End If
        Else
            MkDir strBackupPath
        End If
    End If
    
    'Initialize Decoration
    With frmMain
        If .optDecor(1).Value = True Then
            strMajorDecor = "="
            strMinorDecor = "-"
        ElseIf .optDecor(2).Value = True Then
            strMajorDecor = """"
            strMinorDecor = "'"
        ElseIf .optDecor(3).Value = True Then
            strMajorDecor = "*"
            strMinorDecor = "."
        End If
    End With
    
    For bytCount = 1 To bytModuleSize
        With modpModules(bytCount)
            If .FileName <> "" Then
                'Pre-Step - Initialize Variables & Backup
                ReDim strCode(1 To 1) As String
                intLine = 0
                CopyFile .FileName, strBackupPath & "\" & Right$(.FileName, Len(.FileName) - LastStr(1, .FileName, "\")), &O0
                
                'Step 3.1 - Load code into memory
                Open .FileName For Input As #1
                    Do While Not EOF(1)
                        intLine = intLine + 1
                        ReDim Preserve strCode(1 To intLine) As String
                        Line Input #1, strCode(intLine)
                    Loop
                Close
                
                'Step 3.2 - Find the code line & comment basic info
                Open .FileName For Output As #1
                    'What to do when we've got the file?
                    '3.2.1 Search for the 'Keyword' that tell us the next line is the code
                    '3.2.2 Check for the specified options
                    '3.2.3 Comment the module
                    
                    '3.2.1
                    Select Case .ModType
                    Case MForm
                        'Keyword = "Attribute VB_Exposed" (20 Char)
                        For intLine = 1 To UBound(strCode)
                            If Left$(strCode(intLine), 20) = "Attribute VB_Exposed" Then
                                intLineToIns = intLine + 1
                                Exit For
                            End If
                        Next intLine
                        strModuleType = "Form"
                    Case MModule
                        'No Keyword, There's only 1 line before the code
                        intLineToIns = 2
                        strModuleType = "Form"
                    Case MClass
                        'Keyword = "Attribute VB_Exposed" (20 Char)
                        For intLine = 1 To UBound(strCode)
                            If Left$(strCode(intLine), 20) = "Attribute VB_Exposed" Then
                                intLineToIns = intLine + 1
                                Exit For
                            End If
                        Next intLine
                        strModuleType = "Class"
                    Case MUserControl
                        'Keyword = "Attribute VB_Exposed" (20 Char)
                        For intLine = 1 To UBound(strCode)
                            If Left$(strCode(intLine), 20) = "Attribute VB_Exposed" Then
                                intLineToIns = intLine + 1
                                Exit For
                            End If
                        Next intLine
                        strModuleType = "UserControl"
                    Case MPropPage
                        'Keyword = "Attribute VB_Exposed" (20 Char)
                        For intLine = 1 To UBound(strCode)
                            If Left$(strCode(intLine), 20) = "Attribute VB_Exposed" Then
                                intLineToIns = intLine + 1
                                Exit For
                            End If
                        Next intLine
                        strModuleType = "Property Page"
                    Case MProject
                        intLineToIns = 5
                    End Select
                    
                    '3.2.2 & 3.2.3
                    For intLine = 1 To (intLineToIns - 1)
                        Print #1, strCode(intLine)
                    Next intLine
                    If .ModType <> MProject Then
                        With frmMain
                            booCommentOne = False
                            For bytCheck = 2 To 7
                                If .chkComment(bytCheck).Value = 1 Then booCommentOne = True
                            Next bytCheck
                            If booCommentOne Then
                                'Decoration 1st
                                Print #1,
                                Print #1, "'   " & Multiple(strMajorDecor, 60) '====================
                                Print #1, "'    " & Multiple(strMinorDecor, 58) '------------------
                                If .chkComment(5).Value = 1 Then Print #1, "'     Application Name: " & .txtProjInfo(1)
                                If .chkComment(6).Value = 1 Then Print #1, "'                       " & .txtProjInfo(2)
                                If .chkComment(3).Value = 1 Then Print #1, "'     Developer/Programmer: " & .txtProjInfo(3)
                                Print #1, "'    " & Multiple(strMinorDecor, 58) '------------------
                                If .chkComment(2).Value = 1 Then
                                    Print #1, "'     Module Name: " & modpModules(bytCount).ModName
                                    Print #1, "'     Module File: " & modpModules(bytCount).FileName
                                    Print #1, "'     Module Type: " & strModuleType
                                End If
                                If .chkComment(7).Value = 1 Then Print #1, "'     Module Description:"
                                Print #1, "'    " & Multiple(strMinorDecor, 58) '------------------
                                If .chkComment(4).Value = 1 Then Print #1, "'     " & .txtProjInfo(4)
                                Print #1, "'    " & Multiple(strMinorDecor, 58) '------------------
                                Print #1, "'   " & Multiple(strMajorDecor, 60) '====================
                                Print #1,
                            End If
                        End With
                    Else
                        With frmMain
                            If .chkComment(1).Value = 1 Then
                                '3.2.3.1 Create the info file first
                                Open Left$(strTargetProject, LastStr(1, strTargetProject, "\")) & frmMain.txtProjInfo(1) & " Info.txt" For Output As #2
                                    Print #2, "Project info by Orion Belt " & App.Major & "." & App.Minor
                                    Print #2, "Best viewed in Notepad with a fixed-width font, such as Courier New and Terminal"
                                    Print #2,
                                    Print #2, Multiple(strMajorDecor, 70)
                                    Print #2, Multiple(strMinorDecor, 70)
                                    Print #2, .txtProjInfo(1)
                                    Print #2, .txtProjInfo(2)
                                    Print #2, "By " & .txtProjInfo(3)
                                    If .txtProjInfo(5) <> "" Then
                                        Print #2, Multiple(strMinorDecor, 70)
                                        Print #2, .txtProjInfo(5)
                                    End If
                                    If .txtProjInfo(4) <> "" Then
                                        Print #2, Multiple(strMinorDecor, 70)
                                        Print #2, .txtProjInfo(4)
                                    End If
                                    Print #2, Multiple(strMinorDecor, 70)
                                    Print #2, Multiple(strMajorDecor, 70)
                                    Print #2,
                                    Print #2, "Commented by Orion Belt " & App.Major & "." & App.Minor
                                Close #2
                                '3.2.3.2 Then add a link to the file
                                Print #1, "RelatedDoc=" & frmMain.txtProjInfo(1) & " Info.txt"
                            End If
                        End With
                    End If
                    
                    'Step 3.3 - Comment each procedures according to the options chosen
                    'And what should we do now?
                    '3.3.1 Loop and search for the starting of each procedure
                    '3.3.2 Check for the specified options
                    '3.3.3 Comment procedures according to the options & procedure type
                    'Let Us Go!
                    If frmMain.chkComment(8).Value = 1 Or frmMain.chkComment(9).Value = 1 Then
                        For intLine = intLineToIns To UBound(strCode)
                            bytChar = 0
                            strCurLine = LTrim$(strCode(intLine)) 'In case there're some spaces before the actual code
                            If Left$(strCurLine, 3) = "Sub" Then bytChar = 5
                            If Left$(strCurLine, 8) = "Function" Then bytChar = 10
                            If Left$(strCurLine, 8) = "Property" Then bytChar = 10
                            If Left$(strCurLine, 10) = "Public Sub" Then bytChar = 12
                            If Left$(strCurLine, 11) = "Private Sub" Then bytChar = 13
                            If Left$(strCurLine, 12) = "Public Static" Then bytChar = 14
                            If Left$(strCurLine, 15) = "Public Function" Then bytChar = 17
                            If Left$(strCurLine, 15) = "Public Property" Then bytChar = 17
                            If Left$(strCurLine, 16) = "Private Function" Then bytChar = 18
                            If Left$(strCurLine, 16) = "Private Property" Then bytChar = 18
                            If bytChar > 0 Then
                                If Trim$(strCode(intLine - 1)) = "" Then Print #1,
                                Print #1, "'" & Multiple(strMinorDecor, 40)
                                If frmMain.chkComment(8).Value = 1 Then
                                    Print #1, "'Name: " & Mid$(strCurLine, bytChar, InStr(bytChar, strCurLine, "(") - bytChar)
                                    bytSeperator = InStr(1, strCurLine, "_")
                                    If bytSeperator <> 0 Then 'Object-Based Procedures
                                        Print #1, "'Object: " & Mid$(strCurLine, bytChar, bytSeperator - bytChar)
                                        Print #1, "'Event: " & Mid$(strCurLine, bytSeperator + 1, InStr(bytChar, strCurLine, "(") - (bytSeperator + 1))
                                    End If
                                End If
                                If frmMain.chkComment(9).Value = 1 Then Print #1, "'Description: "
                                Print #1, "'" & Multiple(strMinorDecor, 40)
                            End If
                            Print #1, strCode(intLine)
                        Next intLine
                    End If
                    'Ending Credits
                    If .ModType <> MProject Then
                        Print #1,
                        Print #1, "'Fully commented by Orion Belt®"
                        Print #1, "'©opyright 2001-2002 by Alph@ - All Right Reserved"
                    End If
                Close
            End If
        End With
    Next bytCount
End Sub


'----------------------------------------
'Name: Multiple
'----------------------------------------
Private Function Multiple(strInput As String, bytAmount As Byte) As String
    Dim bytCount As Byte
    For bytCount = 1 To bytAmount
        Multiple = Multiple + strInput
    Next bytCount
End Function


'----------------------------------------
'Name: LastStr
'----------------------------------------
Private Function LastStr(Start As String, String1 As String, String2 As String) As Byte
    Dim bytChar As Integer
    Dim bytStringPos As Byte
    For bytChar = Start To Len(String1)
        bytStringPos = InStr(bytChar, String1, String2)
        If bytStringPos <> 0 Then LastStr = bytStringPos
    Next bytChar
End Function


'----------------------------------------
'Name: Form_Load
'Object: Form
'Event: Load
'----------------------------------------
Private Sub Form_Load()
    Me.Show
End Sub

'Fully commented by Orion Belt®
'©opyright 2001-2002 by Alph@ - All Right Reserved
