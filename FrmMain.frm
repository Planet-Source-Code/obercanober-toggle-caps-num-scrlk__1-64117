VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CNS Toggle"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdNum 
      Caption         =   "Num Lock"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton CmdScroll 
      Caption         =   "Scroll Lock"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton CmdCaps 
      Caption         =   "Caps Lock"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.Timer TmrChkStatus 
      Interval        =   100
      Left            =   2040
      Top             =   480
   End
   Begin VB.CheckBox ChkNum 
      Caption         =   "Num Lock"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox ChkScroll 
      Caption         =   "Scroll Lock"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CheckBox ChkCaps 
      Caption         =   "Caps Lock"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCaps_Click()
    ToggleCapsLock
End Sub

Private Sub CmdNum_Click()
    ToggleNumLock
End Sub

Private Sub CmdScroll_Click()
    ToggleScrollLock
End Sub

Private Sub TmrChkStatus_Timer()
    If CapsLockOn = True Then Me.ChkCaps.Value = 1 Else Me.ChkCaps.Value = 0
    If NumLockOn = True Then Me.ChkNum.Value = 1 Else Me.ChkNum.Value = 0
    If ScrlLockOn = True Then Me.ChkScroll.Value = 1 Else Me.ChkScroll.Value = 0
End Sub
