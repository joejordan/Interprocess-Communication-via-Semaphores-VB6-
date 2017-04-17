VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Semaphore Test"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   3345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdToggleHandle 
      Caption         =   "Toggle Handle Security"
      Height          =   360
      Left            =   600
      TabIndex        =   8
      Top             =   3000
      Width           =   2190
   End
   Begin VB.CommandButton cmdSemSecured 
      Caption         =   "Is Sem Handle Secured?"
      Height          =   360
      Left            =   600
      TabIndex        =   7
      Top             =   2520
      Width           =   2190
   End
   Begin VB.CommandButton cmdQueryHandle 
      Caption         =   "Query Handle Count"
      Height          =   360
      Left            =   600
      TabIndex        =   6
      Top             =   4620
      Width           =   2190
   End
   Begin VB.CommandButton cmdCheckFor 
      Caption         =   "Semaphore Name Exists?"
      Height          =   360
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Width           =   2190
   End
   Begin VB.CommandButton cmdQueryName 
      Caption         =   "Query Semaphore Name"
      Height          =   360
      Left            =   600
      TabIndex        =   4
      Top             =   4140
      Width           =   2190
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Query Semaphore Value"
      Height          =   360
      Left            =   600
      TabIndex        =   3
      Top             =   3660
      Width           =   2190
   End
   Begin VB.CommandButton cmdDecrement 
      Caption         =   "Decrement Semaphore"
      Height          =   360
      Left            =   600
      TabIndex        =   2
      Top             =   1380
      Width           =   2190
   End
   Begin VB.CommandButton cmdIncrement 
      Caption         =   "Increment Semaphore"
      Height          =   360
      Left            =   600
      TabIndex        =   1
      Top             =   900
      Width           =   2190
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create/Open Semaphore"
      Height          =   360
      Left            =   420
      TabIndex        =   0
      Top             =   420
      Width           =   2625
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objIPC As clsIPC

Private Sub cmdCheckFor_Click()
    Dim bRet As Boolean
    bRet = objIPC.IsSemaphore("TestSemaphore1", True)
    MsgBox "Semaphore Exists = " & bRet
End Sub

Private Sub cmdCreate_Click()
    Dim bRet As Boolean
    bRet = objIPC.Initialize("TestSemaphore1", 0, 100, True, False)
    MsgBox "Create Success: " & bRet & vbNewLine & "Handle Value: " & objIPC.SemaphoreHandle
End Sub

Private Sub cmdDecrement_Click()
    Dim bRet As Boolean
    bRet = objIPC.Decrement
    MsgBox "Decrement Success: " & bRet
End Sub

Private Sub cmdIncrement_Click()
    Dim bRet As Boolean, lPrevVal As Long
    bRet = objIPC.Increment(1, lPrevVal)
    MsgBox "Increment Success: " & bRet & vbNewLine & "Previous Value: " & lPrevVal
End Sub

Private Sub cmdQuery_Click()
    Dim lRet As Long, lMax As Long
    lRet = objIPC.QueryCurrentValue
    lMax = objIPC.QueryMaxValue
    MsgBox "Current Value: " & lRet & vbNewLine & "Max Value: " & lMax
End Sub

Private Sub cmdQueryHandle_Click()
    Dim lRet As Long
    lRet = objIPC.QueryHandleCount
    MsgBox "Semaphore Open Handle Count: " & lRet
End Sub

Private Sub cmdQueryName_Click()
    MsgBox objIPC.QueryName
End Sub

Private Sub cmdSemSecured_Click()
    Dim bRet As Boolean
    MsgBox "Semaphore Handle: " & objIPC.SemaphoreHandle & vbNewLine & "Handle secured from closing: " & objIPC.SemaphoreSecurityEnabled
End Sub

Private Sub cmdToggleHandle_Click()
    Dim bRet As Boolean
    bRet = objIPC.SemaphoreSecurityEnabled
    objIPC.SemaphoreSecurityEnabled = Not bRet
    MsgBox "Semaphore Security Enabled: " & objIPC.SemaphoreSecurityEnabled
End Sub

Private Sub Form_Load()
    Set objIPC = New clsIPC
End Sub
