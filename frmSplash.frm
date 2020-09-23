VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1995
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   5100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1950
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4980
      Begin VB.Timer tmrSplash 
         Interval        =   500
         Left            =   75
         Top             =   150
      End
      Begin VB.Image imgLogo 
         Height          =   1320
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   75
         Width           =   1215
      End
      Begin VB.Label lblWarning 
         Caption         =   " Warning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   150
         TabIndex        =   1
         Top             =   1200
         Width           =   2145
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1935
         TabIndex        =   2
         Top             =   825
         Width           =   2565
      End
      Begin VB.Label lblOperSysInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Platform"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2310
         TabIndex        =   3
         Top             =   1425
         Width           =   2535
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1935
         TabIndex        =   4
         Top             =   300
         Width           =   2610
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim objOperSys As New cOperSystem  ' Define and instantiate classes
    
    With frmSplash
        .lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
        .lblProductName.Caption = PGM_NAME
        .lblWarning.Caption = "This is a freeware product." & vbNewLine & _
                              "No warranties or guarantees implied or intended."
        
        ' Capture information about this operating system
        .lblOperSysInfo.Caption = objOperSys.VersionName & vbNewLine & _
                                  "Version " & objOperSys.VersionNumber & "  " & _
                                  objOperSys.ServicePack
        .tmrSplash.Enabled = True
        
        ' center form on screen
        .Move (Screen.Width - frmSplash.Width) \ 2, (Screen.Height - frmSplash.Height) \ 2
        .Show
    End With
    
    Set objOperSys = Nothing  ' Free class objects form memory
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Unload frmSplash          ' Deactivate this form
    Set frmSplash = Nothing   ' Remove form from memory

End Sub

Private Sub tmrSplash_Timer()

    ' gblnFormsLoaded is set to FALSE when application is
    ' first started.  When last form is finished loading,
    ' flag is set to TRUE.
    If gblnFormsLoaded Then
        tmrSplash.Enabled = False  ' Turn off timer
        frmMain.ShowMainForm       ' Display main form
        Unload Me                  ' Unload this form
    End If
    
End Sub


