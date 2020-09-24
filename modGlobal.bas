Attribute VB_Name = "modGlobal"

'   ============================================================
'    ----------------------------------------------------------
'     Application Name: Orion Belt
'                       The World's Best VB Code Commenter
'     Developer/Programmer: Alph@
'    ----------------------------------------------------------
'     Module Name: modGlobal
'     Module File: modGlobal.bas
'     Module Type: Form
'     Module Description:
'    ----------------------------------------------------------
'     ©opyright 2001-2002 by Alph@ - All Right Reserved
'    ----------------------------------------------------------
'   ============================================================

Option Explicit

'Global Declarations
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Enum ModuleType
    MProject = 0
    MForm = 1
    MModule = 2
    MClass = 3
    MUserControl = 4
    MPropPage = 5
End Enum

Type ModuleProperties
    ModName As String
    FileName As String
    ModType As ModuleType
End Type

Public strTargetProject As String

'Fully commented by Orion Belt®
'©opyright 2001-2002 by Alph@ - All Right Reserved
