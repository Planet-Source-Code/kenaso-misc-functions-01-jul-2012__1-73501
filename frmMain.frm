VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7815
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   8295
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0442
   ScaleHeight     =   7815
   ScaleWidth      =   8295
   Begin VB.PictureBox Picture1 
      Height          =   7860
      Left            =   0
      ScaleHeight     =   7800
      ScaleWidth      =   8250
      TabIndex        =   35
      Top             =   0
      Width           =   8310
      Begin VB.PictureBox picTitle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   7605
         Picture         =   "frmMain.frx":0884
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   97
         Top             =   240
         Width           =   480
      End
      Begin VB.CommandButton cmdChoice 
         BackColor       =   &H00E0E0E0&
         Height          =   650
         Index           =   1
         Left            =   7455
         Picture         =   "frmMain.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Terminate this application"
         Top             =   7065
         Width           =   650
      End
      Begin VB.CommandButton cmdChoice 
         BackColor       =   &H00E0E0E0&
         Height          =   650
         Index           =   0
         Left            =   6780
         Picture         =   "frmMain.frx":0FD0
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Display About Screen"
         Top             =   7065
         Width           =   650
      End
      Begin VB.Timer tmrDate 
         Interval        =   1000
         Left            =   4875
         Top             =   7215
      End
      Begin TabDlg.SSTab tabMain 
         Height          =   6000
         Left            =   60
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1020
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   10583
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Dates"
         TabPicture(0)   =   "frmMain.frx":12DA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraDemo(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Random && Sorting"
         TabPicture(1)   =   "frmMain.frx":12F6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraDemo(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Manipulating Bits"
         TabPicture(2)   =   "frmMain.frx":1312
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraDemo(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Drive Info"
         TabPicture(3)   =   "frmMain.frx":132E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "fraDemo(3)"
         Tab(3).ControlCount=   1
         Begin VB.Frame fraDemo 
            BorderStyle     =   0  'None
            Height          =   5595
            Index           =   3
            Left            =   -74955
            TabIndex        =   98
            Top             =   315
            Width           =   7935
            Begin VB.CommandButton cmdRefresh 
               Height          =   330
               Left            =   180
               Picture         =   "frmMain.frx":134A
               Style           =   1  'Graphical
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   1530
               Width           =   465
            End
            Begin VB.ComboBox cboDrive 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   135
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   315
               Width           =   960
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Bytes per sector"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   11
               Left            =   0
               TabIndex        =   142
               Top             =   3960
               Width           =   1395
            End
            Begin VB.Label lblDrvData 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   1485
               TabIndex        =   141
               Top             =   3930
               Width           =   1275
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Formatted drive size"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   18
               Left            =   4680
               TabIndex        =   33
               Top             =   3555
               Width           =   1845
            End
            Begin VB.Label lblDrvData 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   18
               Left            =   6615
               TabIndex        =   140
               Top             =   3555
               Width           =   1275
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Drive Information"
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   675
               TabIndex        =   139
               Top             =   4635
               Width           =   1650
            End
            Begin VB.Label lblDiskTitle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Disk Information"
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   720
               TabIndex        =   138
               Top             =   2145
               Width           =   1650
            End
            Begin VB.Label lblDrive 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Refresh drive selection"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Index           =   1
               Left            =   135
               TabIndex        =   137
               Top             =   855
               Width           =   765
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Available space"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   21
               Left            =   4185
               TabIndex        =   38
               Top             =   4410
               Width           =   1260
            End
            Begin VB.Label lblDrvData 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   21
               Left            =   5535
               TabIndex        =   136
               Top             =   4410
               Width           =   2355
            End
            Begin VB.Label lblDrvData 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   22
               Left            =   5535
               TabIndex        =   135
               Top             =   4695
               Width           =   2355
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Free space"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   22
               Left            =   4185
               TabIndex        =   40
               Top             =   4725
               Width           =   1260
            End
            Begin VB.Label lblDrvData 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   20
               Left            =   5535
               TabIndex        =   134
               Top             =   4125
               Width           =   2355
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Used space"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   20
               Left            =   4185
               TabIndex        =   36
               Top             =   4125
               Width           =   1260
            End
            Begin VB.Label lblDrvData 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   19
               Left            =   5535
               TabIndex        =   133
               Top             =   3840
               Width           =   2355
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Total drive space"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   19
               Left            =   4215
               TabIndex        =   34
               Top             =   3870
               Width           =   1230
            End
            Begin VB.Label lblDrvData 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   1485
               TabIndex        =   132
               Top             =   2505
               Width           =   2355
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Total disk space"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   6
               Left            =   135
               TabIndex        =   131
               Top             =   2505
               Width           =   1260
            End
            Begin VB.Label lblDrvData 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   24
               Left            =   5535
               TabIndex        =   130
               Top             =   5265
               Width           =   2355
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Clusters free"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   24
               Left            =   4185
               TabIndex        =   43
               Top             =   5265
               Width           =   1260
            End
            Begin VB.Label lblDrvData 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   23
               Left            =   5535
               TabIndex        =   129
               Top             =   4980
               Width           =   2355
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Total clusters"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   23
               Left            =   4185
               TabIndex        =   41
               Top             =   4980
               Width           =   1260
            End
            Begin VB.Label lblDrvData 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   17
               Left            =   6615
               TabIndex        =   128
               Top             =   3270
               Width           =   1275
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Drive bytes per cluster"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   17
               Left            =   4740
               TabIndex        =   32
               Top             =   3270
               Width           =   1785
            End
            Begin VB.Label lblDrvData 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   16
               Left            =   6615
               TabIndex        =   127
               Top             =   2985
               Width           =   1275
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Drive bytes per sector"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   16
               Left            =   4740
               TabIndex        =   31
               Top             =   2985
               Width           =   1785
            End
            Begin VB.Label lblDrvData 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   6615
               TabIndex        =   126
               Top             =   2700
               Width           =   1275
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Drive sectors per cluster"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   15
               Left            =   4740
               TabIndex        =   30
               Top             =   2730
               Width           =   1785
            End
            Begin VB.Label lblDrvData 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   10
               Left            =   1485
               TabIndex        =   125
               Top             =   3645
               Width           =   1275
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Sectors per track"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   10
               Left            =   0
               TabIndex        =   124
               Top             =   3675
               Width           =   1395
            End
            Begin VB.Label lblDrvData 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   1485
               TabIndex        =   123
               Top             =   3360
               Width           =   1275
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Tracks per cylinder"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   9
               Left            =   0
               TabIndex        =   122
               Top             =   3390
               Width           =   1395
            End
            Begin VB.Label lblDrvData 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   1485
               TabIndex        =   121
               Top             =   3075
               Width           =   1275
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Total cylinders"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   8
               Left            =   0
               TabIndex        =   120
               Top             =   3075
               Width           =   1395
            End
            Begin VB.Label lblDrvData 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   1485
               TabIndex        =   119
               Top             =   2790
               Width           =   1275
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Formatted disk size"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   7
               Left            =   0
               TabIndex        =   118
               Top             =   2790
               Width           =   1395
            End
            Begin VB.Label lblDrvData 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   6615
               TabIndex        =   117
               Top             =   2415
               Width           =   1275
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "File system type"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   14
               Left            =   5415
               TabIndex        =   116
               Top             =   2445
               Width           =   1110
            End
            Begin VB.Label lblDrvData 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   1470
               TabIndex        =   115
               Top             =   5265
               Width           =   1275
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Volume serial"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   13
               Left            =   435
               TabIndex        =   114
               Top             =   5310
               Width           =   930
            End
            Begin VB.Label lblDrvData 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   1470
               TabIndex        =   113
               Top             =   4980
               Width           =   2400
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Volume name"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   12
               Left            =   420
               TabIndex        =   112
               Top             =   5025
               Width           =   960
            End
            Begin VB.Label lblDrvData 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   5
               Left            =   2790
               TabIndex        =   111
               Top             =   1800
               Width           =   5100
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Manufacturer firmware"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   5
               Left            =   855
               TabIndex        =   110
               Top             =   1800
               Width           =   1815
            End
            Begin VB.Label lblDrvData 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   4
               Left            =   2790
               TabIndex        =   109
               Top             =   1485
               Width           =   5100
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Manufacturer serial"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   4
               Left            =   855
               TabIndex        =   108
               Top             =   1485
               Width           =   1815
            End
            Begin VB.Label lblDrvData 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   3
               Left            =   2790
               TabIndex        =   107
               Top             =   1170
               Width           =   5100
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Manufacturer model"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   3
               Left            =   855
               TabIndex        =   106
               Top             =   1170
               Width           =   1815
            End
            Begin VB.Label lblDrvData 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   2
               Left            =   2790
               TabIndex        =   105
               Top             =   855
               Width           =   5100
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Partition "
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   2
               Left            =   1440
               TabIndex        =   104
               Top             =   855
               Width           =   1275
            End
            Begin VB.Label lblDrvData 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   1
               Left            =   2790
               TabIndex        =   103
               Top             =   540
               Width           =   5100
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Drive type extra"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   1
               Left            =   1440
               TabIndex        =   102
               Top             =   540
               Width           =   1275
            End
            Begin VB.Label lblDrvInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Drive type"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   0
               Left            =   1440
               TabIndex        =   101
               Top             =   225
               Width           =   1275
            End
            Begin VB.Label lblDrvData 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   0
               Left            =   2790
               TabIndex        =   100
               Top             =   225
               Width           =   5100
            End
            Begin VB.Label lblDrive 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Select drive"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   180
               TabIndex        =   99
               Top             =   90
               Width           =   900
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame fraDemo 
            BorderStyle     =   0  'None
            Height          =   5535
            Index           =   0
            Left            =   90
            TabIndex        =   68
            Top             =   330
            Width           =   7905
            Begin VB.TextBox txtDateResults 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   13
               Left            =   1710
               Locked          =   -1  'True
               TabIndex        =   152
               TabStop         =   0   'False
               Top             =   2460
               Width           =   3015
            End
            Begin VB.TextBox txtDateResults 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   12
               Left            =   6675
               Locked          =   -1  'True
               TabIndex        =   146
               TabStop         =   0   'False
               Top             =   2430
               Width           =   1125
            End
            Begin VB.TextBox txtDateResults 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   11
               Left            =   6675
               Locked          =   -1  'True
               TabIndex        =   145
               TabStop         =   0   'False
               Top             =   1965
               Width           =   1125
            End
            Begin VB.ComboBox cboDayOfMonth 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1710
               TabIndex        =   1
               TabStop         =   0   'False
               Text            =   "cboDayOfMonth"
               Top             =   1965
               Width           =   3495
            End
            Begin VB.TextBox txtDateResults 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Index           =   8
               Left            =   1710
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   79
               TabStop         =   0   'False
               Top             =   3820
               Width           =   6090
            End
            Begin VB.TextBox txtDateResults 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   10
               Left            =   5505
               Locked          =   -1  'True
               TabIndex        =   78
               TabStop         =   0   'False
               Top             =   5160
               Width           =   1650
            End
            Begin VB.TextBox txtDateResults 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   9
               Left            =   3735
               Locked          =   -1  'True
               TabIndex        =   77
               TabStop         =   0   'False
               Top             =   5160
               Width           =   1650
            End
            Begin VB.TextBox txtDateResults 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   7
               Left            =   1710
               Locked          =   -1  'True
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   4470
               Width           =   2250
            End
            Begin VB.TextBox txtDate 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   135
               TabIndex        =   0
               Top             =   285
               Width           =   7665
            End
            Begin VB.TextBox txtDateResults 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   1710
               Locked          =   -1  'True
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   765
               Width           =   1050
            End
            Begin VB.TextBox txtDateResults 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   4575
               Locked          =   -1  'True
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   765
               Width           =   3225
            End
            Begin VB.TextBox txtDateResults 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   1710
               Locked          =   -1  'True
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   1140
               Width           =   1050
            End
            Begin VB.TextBox txtDateResults 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   3
               Left            =   5550
               Locked          =   -1  'True
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   1140
               Width           =   2250
            End
            Begin VB.TextBox txtDateResults 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   4
               Left            =   1710
               Locked          =   -1  'True
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1050
            End
            Begin VB.TextBox txtDateResults 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   5
               Left            =   5550
               Locked          =   -1  'True
               TabIndex        =   70
               TabStop         =   0   'False
               Top             =   1515
               Width           =   2250
            End
            Begin VB.TextBox txtDateResults 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Index           =   6
               Left            =   1710
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   3170
               Width           =   6090
            End
            Begin VB.Label lblDate 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Daylight Savings"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   16
               Left            =   195
               TabIndex        =   151
               Top             =   2400
               Width           =   1425
            End
            Begin VB.Label lblDate 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   510
               Index           =   15
               Left            =   1980
               TabIndex        =   148
               Top             =   5085
               Width           =   1605
            End
            Begin VB.Label lblDate 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Easter Sunday"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   14
               Left            =   5325
               TabIndex        =   147
               Top             =   2505
               Width           =   1260
            End
            Begin VB.Label lblDate 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Day of month"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   13
               Left            =   5415
               TabIndex        =   144
               Top             =   2025
               Width           =   1170
            End
            Begin VB.Label lblDate 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Sample dates for "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   12
               Left            =   195
               TabIndex        =   143
               Top             =   1905
               Width           =   1425
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblDate 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Time to words"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   9
               Left            =   90
               TabIndex        =   92
               Top             =   3945
               Width           =   1545
            End
            Begin VB.Label lblY2K38 
               Appearance      =   0  'Flat
               Caption         =   """The next Y2K"" article by Karl E. Peterson"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   525
               Left            =   60
               TabIndex        =   91
               Top             =   5085
               Width           =   1815
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblDate 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Seconds to date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   11
               Left            =   5595
               TabIndex        =   90
               Top             =   4905
               Width           =   1560
            End
            Begin VB.Label lblDate 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Date to seconds"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   10
               Left            =   4005
               TabIndex        =   89
               Top             =   4905
               Width           =   1185
            End
            Begin VB.Label lblDate 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Current time"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   8
               Left            =   90
               TabIndex        =   88
               Top             =   4545
               Width           =   1545
            End
            Begin VB.Label lblDate 
               BackStyle       =   0  'Transparent
               Caption         =   "Enter date  (Any valid format)   Press TAB key"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   165
               TabIndex        =   87
               Top             =   60
               Width           =   4545
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblDate 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Short format"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   90
               TabIndex        =   86
               Top             =   825
               Width           =   1545
            End
            Begin VB.Label lblDate 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Long format"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   3405
               TabIndex        =   85
               Top             =   825
               Width           =   1095
            End
            Begin VB.Label lblDate 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date to Julian"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   90
               TabIndex        =   84
               Top             =   1185
               Width           =   1545
            End
            Begin VB.Label lblDate 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Julian to date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   4
               Left            =   3780
               TabIndex        =   83
               Top             =   1185
               Width           =   1695
            End
            Begin VB.Label lblDate 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date to serial"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   5
               Left            =   90
               TabIndex        =   82
               Top             =   1560
               Width           =   1545
            End
            Begin VB.Label lblDate 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Serial to date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   6
               Left            =   3780
               TabIndex        =   81
               Top             =   1560
               Width           =   1695
            End
            Begin VB.Label lblDate 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date to words"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   7
               Left            =   90
               TabIndex        =   80
               Top             =   3315
               Width           =   1545
            End
         End
         Begin VB.Frame fraDemo 
            BorderStyle     =   0  'None
            Height          =   5415
            Index           =   2
            Left            =   -74865
            TabIndex        =   54
            Top             =   450
            Width           =   7305
            Begin VB.TextBox txtConvert 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   2
               Left            =   75
               MultiLine       =   -1  'True
               TabIndex        =   22
               Text            =   "frmMain.frx":144C
               Top             =   4200
               Width           =   7065
            End
            Begin VB.TextBox txtConvert 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   1
               Left            =   1440
               TabIndex        =   21
               Text            =   "txtConvert(1)"
               Top             =   3585
               Width           =   5700
            End
            Begin VB.TextBox txtConvert 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   0
               Left            =   1440
               TabIndex        =   20
               Text            =   "txtConvert(0)"
               Top             =   3150
               Width           =   5700
            End
            Begin VB.CommandButton cmdConvert 
               Caption         =   "Clear &boxes"
               Height          =   480
               Index           =   1
               Left            =   4350
               TabIndex        =   23
               Top             =   4800
               Width           =   1170
            End
            Begin VB.CommandButton cmdConvert 
               Caption         =   "&Convert"
               Height          =   480
               Index           =   0
               Left            =   5595
               TabIndex        =   24
               Top             =   4800
               Width           =   1170
            End
            Begin VB.PictureBox Picture2 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   420
               Left            =   3240
               ScaleHeight     =   420
               ScaleWidth      =   3300
               TabIndex        =   65
               Top             =   450
               Width           =   3300
               Begin VB.OptionButton optValueType 
                  Caption         =   "Long Integer"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   525
                  Index           =   0
                  Left            =   135
                  TabIndex        =   14
                  Top             =   -90
                  Value           =   -1  'True
                  Width           =   870
               End
               Begin VB.OptionButton optValueType 
                  Caption         =   "Short Integer"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   525
                  Index           =   1
                  Left            =   1125
                  TabIndex        =   15
                  Top             =   -90
                  Width           =   870
               End
               Begin VB.OptionButton optValueType 
                  Caption         =   "Byte"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   525
                  Index           =   2
                  Left            =   2205
                  TabIndex        =   16
                  Top             =   -90
                  Width           =   870
               End
            End
            Begin VB.CommandButton cmdBits 
               Caption         =   "&Shift Bits"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   0
               Left            =   4350
               TabIndex        =   18
               Top             =   2100
               Width           =   1170
            End
            Begin VB.TextBox txtBitInput 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   1740
               TabIndex        =   17
               Text            =   "0"
               Top             =   900
               Width           =   1380
            End
            Begin VB.TextBox txtBitOutput 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   3240
               Locked          =   -1  'True
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   900
               Width           =   3540
            End
            Begin VB.TextBox txtBitOutput 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   1740
               Locked          =   -1  'True
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   1260
               Width           =   1380
            End
            Begin VB.TextBox txtBitOutput 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   3225
               Locked          =   -1  'True
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   1260
               Width           =   3540
            End
            Begin VB.TextBox txtBitInput 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   1755
               MaxLength       =   2
               TabIndex        =   13
               Text            =   "0"
               Top             =   450
               Width           =   390
            End
            Begin VB.CommandButton cmdBits 
               Caption         =   "&Rotate Bits"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   1
               Left            =   5595
               TabIndex        =   19
               Top             =   2100
               Width           =   1170
            End
            Begin VB.TextBox txtBitOutput 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   4
               Left            =   3225
               Locked          =   -1  'True
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   1620
               Width           =   3540
            End
            Begin VB.TextBox txtBitOutput 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   3
               Left            =   1740
               Locked          =   -1  'True
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   1620
               Width           =   1380
            End
            Begin VB.Label lblConversion 
               BackStyle       =   0  'Transparent
               Caption         =   "Binary string"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   105
               TabIndex        =   96
               Top             =   3945
               Width           =   1290
            End
            Begin VB.Label lblConversion 
               BackStyle       =   0  'Transparent
               Caption         =   "Hex value"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   105
               TabIndex        =   95
               Top             =   3660
               Width           =   1290
            End
            Begin VB.Label lblConversion 
               BackStyle       =   0  'Transparent
               Caption         =   "Whole number"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   105
               TabIndex        =   94
               Top             =   3225
               Width           =   1290
            End
            Begin VB.Label lblConversion 
               BackStyle       =   0  'Transparent
               Caption         =   "Enter a number or hex value and click convert button"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   93
               Top             =   2835
               Width           =   6270
            End
            Begin VB.Label lblBits 
               BackStyle       =   0  'Transparent
               Caption         =   "Data type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   5
               Left            =   2340
               TabIndex        =   64
               Top             =   525
               Width           =   765
            End
            Begin VB.Label lblBits 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Whole number"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   390
               TabIndex        =   63
               Top             =   975
               Width           =   1230
            End
            Begin VB.Label lblBits 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Bit positions"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   -255
               TabIndex        =   62
               Top             =   525
               Width           =   1875
            End
            Begin VB.Label lblBits 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Left"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   2
               Left            =   960
               TabIndex        =   61
               Top             =   1320
               Width           =   660
            End
            Begin VB.Label lblBits 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Right"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   960
               TabIndex        =   60
               Top             =   1680
               Width           =   660
            End
            Begin VB.Line Line1 
               X1              =   585
               X2              =   6705
               Y1              =   2640
               Y2              =   2640
            End
         End
         Begin VB.Frame fraDemo 
            Height          =   5475
            Index           =   1
            Left            =   -74880
            TabIndex        =   45
            Top             =   390
            Width           =   7890
            Begin VB.CommandButton cmdRnd 
               Caption         =   "&Sort"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   1
               Left            =   6945
               TabIndex        =   12
               Top             =   975
               Width           =   795
            End
            Begin VB.CommandButton cmdRnd 
               Caption         =   "&Create"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   0
               Left            =   6960
               TabIndex        =   11
               Top             =   240
               Width           =   795
            End
            Begin VB.Frame fraDataType 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   1350
               Left            =   120
               TabIndex        =   51
               Top             =   120
               Width           =   3810
               Begin VB.PictureBox picDateFmt 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   1095
                  Left            =   1020
                  ScaleHeight     =   1095
                  ScaleWidth      =   2700
                  TabIndex        =   154
                  Top             =   180
                  Width           =   2700
                  Begin VB.ComboBox cboFormat 
                     BeginProperty Font 
                        Name            =   "Courier New"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Index           =   1
                     Left            =   60
                     TabIndex        =   156
                     Top             =   720
                     Width           =   2655
                  End
                  Begin VB.ComboBox cboFormat 
                     BeginProperty Font 
                        Name            =   "Courier New"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Index           =   0
                     Left            =   60
                     TabIndex        =   6
                     Top             =   720
                     Width           =   2655
                  End
                  Begin VB.ComboBox cboDate 
                     BeginProperty Font 
                        Name            =   "Courier New"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Left            =   60
                     TabIndex        =   5
                     Top             =   60
                     Width           =   2655
                  End
                  Begin VB.Label lblDateFmt 
                     Caption         =   "Date/Time Format"
                     Height          =   195
                     Left            =   60
                     TabIndex        =   155
                     Top             =   480
                     Width           =   1695
                  End
               End
               Begin VB.PictureBox Picture3 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   1035
                  Left            =   105
                  ScaleHeight     =   1035
                  ScaleWidth      =   840
                  TabIndex        =   66
                  Top             =   180
                  Width           =   840
                  Begin VB.OptionButton optDataType 
                     Caption         =   "Date"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   2
                     Left            =   0
                     TabIndex        =   4
                     Top             =   780
                     Width           =   885
                  End
                  Begin VB.OptionButton optDataType 
                     Caption         =   "String"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   1
                     Left            =   0
                     TabIndex        =   3
                     Top             =   450
                     Width           =   885
                  End
                  Begin VB.OptionButton optDataType 
                     Caption         =   "Numeric"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   0
                     Left            =   15
                     TabIndex        =   2
                     Top             =   120
                     Value           =   -1  'True
                     Width           =   885
                  End
               End
            End
            Begin VB.Frame fraSortOrder 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1350
               Left            =   4020
               TabIndex        =   50
               Top             =   120
               Width           =   2850
               Begin VB.PictureBox Picture4 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   720
                  Index           =   0
                  Left            =   45
                  ScaleHeight     =   720
                  ScaleWidth      =   2685
                  TabIndex        =   67
                  Top             =   510
                  Width           =   2685
                  Begin VB.CheckBox chkCaseSensitive 
                     Caption         =   "Case sensitive"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   1320
                     TabIndex        =   10
                     TabStop         =   0   'False
                     Top             =   375
                     Width           =   1395
                  End
                  Begin VB.CheckBox chkDuplicates 
                     Caption         =   "Remove dupes"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   1320
                     TabIndex        =   9
                     TabStop         =   0   'False
                     Top             =   60
                     Width           =   1395
                  End
                  Begin VB.OptionButton optSortOrder 
                     Caption         =   "Ascending"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   0
                     Left            =   30
                     TabIndex        =   7
                     Top             =   60
                     Value           =   -1  'True
                     Width           =   1155
                  End
                  Begin VB.OptionButton optSortOrder 
                     Caption         =   "Descending"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   1
                     Left            =   30
                     TabIndex        =   8
                     Top             =   360
                     Width           =   1155
                  End
               End
               Begin VB.Label lblRnd 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sort options"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   7
                  Left            =   180
                  TabIndex        =   153
                  Top             =   240
                  Width           =   2490
                  WordWrap        =   -1  'True
               End
            End
            Begin VB.ListBox lstRnd 
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2760
               Index           =   1
               Left            =   3975
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   2340
               Width           =   3735
            End
            Begin VB.ListBox lstRnd 
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2760
               Index           =   0
               Left            =   120
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   2340
               Width           =   3735
            End
            Begin VB.Label lblRnd 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Sort method"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   225
               Index           =   3
               Left            =   120
               TabIndex        =   150
               Top             =   1800
               Width           =   3690
            End
            Begin VB.Label lblRnd 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Elapsed time"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   2
               Left            =   4020
               TabIndex        =   149
               Top             =   1560
               Width           =   3690
            End
            Begin VB.Label lblRnd 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Sorted results"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   225
               Index           =   5
               Left            =   4020
               TabIndex        =   53
               Top             =   1800
               Width           =   3690
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblRnd 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Unsorted qty"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   4
               Left            =   120
               TabIndex        =   52
               Top             =   1560
               Width           =   3690
            End
            Begin VB.Label lblRnd 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Sorted"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   4095
               TabIndex        =   49
               Top             =   2100
               Width           =   3495
            End
            Begin VB.Label lblRnd 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Unsorted"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   47
               Top             =   2100
               Width           =   3495
            End
         End
      End
      Begin VB.PictureBox picTitle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   210
         Picture         =   "frmMain.frx":145A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   37
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblOperSys 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   44
         Top             =   7215
         Width           =   3945
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAuthor 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kenneth Ives"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   3660
         TabIndex        =   42
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Caption         =   "kiMisc Dll Demo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2430
         TabIndex        =   39
         Top             =   90
         Width           =   3435
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Routine:       frmMain
'
' Description:   Demonstration of referencing kiMisc.dll
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 27-JUL-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 10-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Added drive information tab
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module constants
' ***************************************************************************
  Private Const MAX_INT       As Long = &H7FFF       '  32767
  Private Const MIN_INT       As Long = &H8000       ' -32768
  Private Const MAX_LONG      As Long = &H7FFFFFFF   '  2147483647
  Private Const MIN_LONG      As Long = &H80000000   ' -2147483648
  Private Const SW_SHOWNORMAL As Long = 1
  Private Const SORT_ITEMS    As Long = 1000

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' Reduce flicker while loading a control
  ' Lock the control to prevent redrawing (flickering)
  '     Syntax:  LockWindowUpdate ctl_name.hWnd
  ' Unlock the control
  '     Syntax:  LockWindowUpdate 0&
  Private Declare Function LockWindowUpdate Lib "user32" _
          (ByVal hwnd As Long) As Long
  
  ' The GetDesktopWindow function returns a handle to the desktop window.
  ' The desktop window covers the entire screen. The desktop window is
  ' the area on top of which other windows are painted.
  Private Declare Function GetDesktopWindow Lib "user32" () As Long

  ' The ShellExecute function opens or prints a specified file.  The file
  ' can be an executable file or a document file.
  Private Declare Function ShellExecute Lib "shell32.dll" _
          Alias "ShellExecuteA" (ByVal hwnd As Long, _
          ByVal lpOperation As String, ByVal lpFile As String, _
          ByVal lpParameters As String, ByVal lpDirectory As String, _
          ByVal nShowCmd As Long) As Long

' ***************************************************************************
' Module variables
' ***************************************************************************
  Private mobjKeyEdit     As cKeyEdit
  Private mobjDates       As kiMisc.cDates
  Private mstrPrevTime    As String
  Private mstrDateFmt     As String
  Private mstrTimeFmt     As String
  Private mlngDateType    As Long
  Private mlngValueType   As Long
  Private mlngSelection   As Long
  Private mlngDateFormat  As Long
  Private mlngTimeFormat  As Long
  Private mdatDateEntered As Date
  
Public Sub ShowMainForm()
        
    With frmMain
        .tmrDate.Enabled = True
        .Show vbModeless
        .Refresh
    End With
    
End Sub

Private Sub LoadComboBoxes()

    ' Called by Form_Load()
    '           cmdRefresh_Click()
    
    Dim lngIndex     As Long
    Dim astrDrives() As String
    Dim objDiskInfo  As kiMisc.cDiskInfo
    
    On Error Resume Next
    
    Set objDiskInfo = New kiMisc.cDiskInfo   ' Instantiate class object
    Erase astrDrives()                       ' Always start with empty arrays
    cboDrive.Clear                           ' Empty combobox
    
    With objDiskInfo
    
        astrDrives() = .GetDriveLetters   ' Capture drive letters in use
        
        For lngIndex = 0 To UBound(astrDrives) - 1
            
            If .SpecificTypeOfDrive(astrDrives(lngIndex), eFixed) Then
                cboDrive.AddItem astrDrives(lngIndex)  ' Add physical hard drive to combobox
            
            ElseIf .SpecificTypeOfDrive(astrDrives(lngIndex), eRemovable, eUSB_Drive) Then
                cboDrive.AddItem astrDrives(lngIndex)  ' Add USB drive to combobox
            End If
            
        Next lngIndex
        
    End With
    
    Set objDiskInfo = Nothing   ' Always free object from memory when not needed
    Erase astrDrives()          ' Always empty arrays when not needed
    cboDrive.ListIndex = 0      ' Set to first item in combo box
    cboDrive.Refresh

    With cboDayOfMonth
        .Clear
        .AddItem "Daylight savings - Start"
        .AddItem "Daylight savings - End"
        .AddItem "Martin L King (3rd Mon Jan)"
        .AddItem "Presidents Day (3rd Mon Feb)"
        .AddItem "Mothers Day (2nd Sun May)"
        .AddItem "Fathers Day (3rd Sun Jun)"
        .AddItem "Thanksgiving (4th Thu Nov)"
        .ListIndex = 0
    End With
    
    With cboDate
        .Clear
        .AddItem "Date and Time"
        .AddItem "Date only"
        .AddItem "Time only"
        .ListIndex = 0
    End With
    
    ' Load date formats
    With cboFormat(0)
        .Clear
        .AddItem "MMM dd, yyyy"    ' 0
        .AddItem "MMM d, yyyy"     ' 1
        .AddItem "MMMM dd, yyyy"   ' 2
        .AddItem "MMMM d, yyyy"    ' 3

        .AddItem "dd-MMM-yyyy"     ' 4
        .AddItem "d-MMM-yyyy"      ' 5
        .AddItem "dd MMM yyyy"     ' 6
        .AddItem "d MMM yyyy"      ' 7
        .AddItem "dd.MMM.yyyy"     ' 8
        .AddItem "d.MMM.yyyy"      ' 9

        .AddItem "yyyy-MMM-dd"     ' 10
        .AddItem "yyyy-MMM-d"      ' 11
        .AddItem "yyyy MMM dd"     ' 12
        .AddItem "yyyy MMM d"      ' 13
        .AddItem "yyyy.MMM.dd"     ' 14
        .AddItem "yyyy.MMM.d"      ' 15

        .AddItem "mm/dd/yyyy"      ' 16
        .AddItem "m/d/yyyy"        ' 17
        .AddItem "mm-dd-yyyy"      ' 18
        .AddItem "m-d-yyyy"        ' 19
        .AddItem "mm.dd.yyyy"      ' 20
        .AddItem "m.d.yyyy"        ' 21

        .AddItem "dd/mm/yyyy"      ' 22
        .AddItem "d/m/yyyy"        ' 23
        .AddItem "dd-mm-yyyy"      ' 24
        .AddItem "d-m-yyyy"        ' 25
        .AddItem "dd.mm.yyyy"      ' 26
        .AddItem "d.m.yyyy"        ' 27

        .AddItem "yyyy/mm/dd"      ' 28
        .AddItem "yyyy/m/d"        ' 29
        .AddItem "yyyy-mm-dd"      ' 30
        .AddItem "yyyy-m-d"        ' 31
        .AddItem "yyyy.mm.dd"      ' 32
        .AddItem "yyyy.m.d"        ' 33

        .AddItem "yyyy/dd/mm"      ' 34
        .AddItem "yyyy/d/m"        ' 35
        .AddItem "yyyy-dd-mm"      ' 36
        .AddItem "yyyy-d-m"        ' 37
        .AddItem "yyyy.dd.mm"      ' 38
        .AddItem "yyyy.d.m"        ' 39
        .ListIndex = 0
    End With
    
    ' Load time formats
    With cboFormat(1)
        .Clear
        .AddItem "h:nn"                     ' 0
        .AddItem "hh:nn"                    ' 1
        .AddItem "hh:nn:ss"                 ' 2
        .AddItem "hh:nn:ss AM/PM"           ' 3
        .AddItem "h:nna/p"                  ' 4
        .AddItem "hh:nnam/pm"               ' 5
                                               
        .AddItem "hh:nn:ss A.M./P.M."       ' 6
        .AddItem "hh:nn:ss.ttt"             ' 7
        .AddItem "hh:nn:ss:ttt"             ' 8
        .AddItem "hh:nn:ss.ttt AM/PM"       ' 9
        .AddItem "hh:nn:ss:ttt AM/PM"       ' 10
        .AddItem "hh:nn:ss.ttt A.M./P.M."   ' 11
        .AddItem "hh:nn:ss:ttt A.M./P.M."   ' 12
                                               
        .AddItem "h.nn"                     ' 13
        .AddItem "hh.nn"                    ' 14
        .AddItem "hh.nn.ss"                 ' 15
        .AddItem "hh.nn.ss AM/PM"           ' 16
        .AddItem "h.nna/p"                  ' 17
        .AddItem "hh.nnam/pm"               ' 18
                                               
        .AddItem "hh.nn.ss A.M./P.M."       ' 19
        .AddItem "hh.nn.ss.ttt"             ' 20
        .AddItem "hh.nn.ss.ttt AM/PM"       ' 21
        .AddItem "hh.nn.ss.ttt A.M./P.M."   ' 22
        
        .AddItem "hh:nn:ss.tttt"            ' 23
        .AddItem "hh:nn:ss:tttt"            ' 24
        .AddItem "hh.nn.ss.tttt"            ' 25
        .AddItem "hh:nn:ss.tttt AM/PM"      ' 26
        .AddItem "hh:nn:ss:tttt AM/PM"      ' 27
        .AddItem "hh.nn.ss.tttt AM/PM"      ' 28
        .AddItem "hh:nn:ss.tttt A.M./P.M."  ' 29
        .AddItem "hh:nn:ss:tttt A.M./P.M."  ' 30
        .AddItem "hh.nn.ss.tttt A.M./P.M."  ' 31
        .ListIndex = 0
    End With
    
    On Error GoTo 0

End Sub
  
Private Sub Random_Demo(ByVal intIndex As Integer)
    
    Dim lngIndex    As Long
    Dim lngCount    As Long
    Dim lngDateTime As Long
    Dim intDay      As Integer
    Dim intYear     As Integer
    Dim intHour     As Integer
    Dim intMonth    As Integer
    Dim intMinute   As Integer
    Dim intSecond   As Integer
    Dim intThousand As Integer
    Dim strRnd      As String
    Dim strFmt1     As String
    Dim strFmt2     As String
    Dim strTemp     As String
    Dim strTime     As String
    Dim strThou     As String
    Dim astrData()  As String
    Dim abytData()  As Byte
    Dim adblData()  As Double
    Dim objRnd      As kiMisc.cPrng
    Dim objSort     As kiMisc.cSort
    
    Const KB_2 As Long = 2048
    Const KB_8 As Long = 8192
    
    strRnd = vbNullString
    strTemp = vbNullString
    strFmt1 = vbNullString
    strFmt2 = vbNullString
    lngCount = 0

    Select Case intIndex
           
           Case 0  ' Create some data
                ' Deactivate so there are no more
                ' changes while data is being created
                fraSortOrder.Enabled = False
                cmdRnd(1).Enabled = False
           
                lstRnd(0).Clear                   ' empty both listboxes
                lstRnd(1).Clear
                lblRnd(2).Caption = vbNullString  ' Clear results
                lblRnd(5).Caption = vbNullString

                Set objRnd = New kiMisc.cPrng     ' Instantiate random class object
                Randomize objRnd.RndSeed          ' Reseed VB Random number generator
                LockWindowUpdate lstRnd(0).hwnd   ' Lock listbox to prevent flickering while loading

                ' Determine data type selected
                Select Case mlngSelection
                       
                       Case 0  ' Numeric data
                            ' 26-Oct-2011 Create optional ways of loading listbox
                            '---------------------------------------------------------
                            
                            Do
                                ' Create string of hex characters using
                                ' MS CryptoAPI random number generator
                                strRnd = objRnd.BuildRndData(KB_2, ePRNG_HEX)

                                ' Parse data and create Long Integers
                                For lngIndex = 1 To (Len(strRnd) - 4) Step 2

                                    ' To make all positive values 0-65535 then
                                    ' append an ampersand "&" to end of data.
                                    ' ex: -1 = &HFFFF   65535 = &HFFFF&
                                    '
                                    ' Create one short integer and add to listbox.
                                    'lstRnd(0).AddItem Val("&H" & Mid$(strRnd, lngIndex, 4))
                                    lstRnd(0).AddItem Val("&H" & Mid$(strRnd, lngIndex, 4))

                                    ' If enough data has been created then exit loop
                                    If lstRnd(0).ListCount = SORT_ITEMS Then
                                        Exit For    ' exit For..Next loop
                                    End If

                                Next lngIndex

                                strRnd = vbNullString   ' Verify variable is empty (save memory resources)

                                ' If enough data has been created then exit loop
                                If lstRnd(0).ListCount = SORT_ITEMS Then
                                    Exit Do    ' exit Do..Loop
                                End If

                            Loop
                                    
                            '*************************************************************
                            ' Generate long integer values using MS CryptoAPI random
                            ' number generator (ex: -32768 to 32767)
                            '*************************************************************
                          '  For lngIndex = 1 To SORT_ITEMS
                          '      lstRnd(0).AddItem objRnd.GetRndValue(MIN_INT, MAX_INT)
                          '  Next lngIndex
                   
                            '*************************************************************
                            ' Generate double precision values using Visual BASIC random
                            ' number generator (ex:  -765.4321 to 765.4320)
                            '*************************************************************
                          '  For lngIndex = 1 To SORT_ITEMS
                          '      lstRnd(0).AddItem (Rnd() * (765.4320 - -765.4321 + 1)) + -765.4321
                          '  Next lngIndex
                   
                       Case 1  ' String data
                            ' 26-Oct-2011  Fixed bug.  Was creating less than
                            '              maximum number of items needed.
                            Do
                                ' Create random data based on letters only
                                Erase abytData()
                                abytData() = objRnd.BuildWithinRange(KB_8, 65, 122)
                                
                                For lngIndex = 0 To UBound(abytData) - 1
                                                                     
                                    ' Is this a valid alphabetic character
                                    Select Case abytData(lngIndex)
                                           Case 65 To 90, 97 To 122  ' Letters only (A-Z, a-z)
                                                strTemp = strTemp & Chr$(abytData(lngIndex))
                                                lngCount = lngCount + 1
                                                
                                                ' Separate first three characters
                                                If lngCount Mod 3 = 0 Then
                                                    strTemp = strTemp & Chr$(32)
                                                End If
                                    End Select
                                    
                                    ' If ten characters have been
                                    ' collected then add to listbox
                                    ' and update counter
                                    If lngCount = 9 Then
                                        lstRnd(0).AddItem Trim$(strTemp)  ' Drop trailing blanks
                                        strTemp = vbNullString
                                        lngCount = 0
                                    
                                        ' If enough data has been created then exit loop
                                        If lstRnd(0).ListCount = SORT_ITEMS Then
                                            Exit For    ' exit For..Next loop
                                        End If
                                    End If
                                    
                                Next lngIndex
                                
                                ' If enough data has been created then exit loop
                                If lstRnd(0).ListCount = SORT_ITEMS Then
                                    Exit Do    ' exit Do..Loop
                                End If
                            
                            Loop
                       
                       Case 2  ' Dates (Numeric sorting)
                            If mlngDateType <= 1 Then
                                If Len(mstrDateFmt) = 0 Then
                                    cboFormat_Click 0   ' Verify date format, if missing
                                End If
                            Else
                                If Len(mstrTimeFmt) = 0 Then
                                    cboFormat_Click 1   ' Verify time format, if missing
                                End If
                            End If
                            
                            Select Case mlngDateFormat
                                   Case 1, 5, 7, 9, 11, 13, 15
                                        strFmt1 = String$(Len(mstrDateFmt) + 1, "@")
                                   
                                   Case 2, 3
                                        strFmt1 = String$(18, "@")
                                   
                                   Case 17, 19, 21, 23, 25, 27, 29, 31, 33, 35, 37, 39
                                        strFmt1 = String$(Len(mstrDateFmt) + 2, "@")
                                   
                                   Case Else
                                        strFmt1 = String$(Len(mstrDateFmt), "@")
                            End Select
                            
                            ' Adjust buffer string lengths
                            Select Case mlngTimeFormat
                                   Case 0, 13
                                        strFmt2 = String$(Len(mstrTimeFmt) + 1, "@")
                
                                   Case 4, 17   ' Drop "p"
                                        strFmt2 = String$(Len(mstrTimeFmt) - 1, "@")
                
                                   Case 3, 5, 9, 10, 16, 18, 21, 26, 27, 28 ' Drop "/PM"
                                        strFmt2 = String$(Len(mstrTimeFmt) - 3, "@")
                                   
                                   Case 6, 11, 12, 19, 22, 29, 30, 31        ' Drop "/P.M."
                                        strFmt2 = String$(Len(mstrTimeFmt) - 5, "@")
                                   
                                   Case Else
                                        strFmt2 = String$(Len(mstrTimeFmt), "@")
                            End Select
                    
                            Select Case mlngTimeFormat
                                   Case 19, 20, 21, 22, 25, 28, 31: strTemp = "."
                                   Case Else:                       strTemp = ":"
                            End Select
                            
                            Select Case mlngTimeFormat
                                   Case 23 To 32: strThou = "0000"
                                   Case Else:     strThou = "000"
                            End Select
                                    
                            For lngCount = 1 To SORT_ITEMS
                            
                                ' I have tested ranges 1600-2199 with no problems.
                                ' If any suggestions, please email me.  Thank you.
                                
                                ' Random date creation
                                intYear = CInt(Int(Rnd() * (2199 - 1900 + 1)) + 1900)   ' Years 1900 to 2199
                                intMonth = CInt(Int(Rnd() * 12) + 1)   ' Months 1-12
                                intDay = CInt(Int(Rnd() * 31) + 1)     ' Days 1-31  If number of days selected exceeds the number
                                                                       ' of days allowed for the selected month then DateSerial()
                                                                       ' will automatically adjust the date correctly for the
                                                                       ' following month.
                                                                       ' ex:  2/31/2012 will become 3/2/2012
                                
                                ' Random time creation
                                intHour = CInt(Int(Rnd() * 24))        ' Hours 0-23 (24-hour clock)
                                intMinute = CInt(Int(Rnd() * 60))      ' Minutes 0-59
                                intSecond = CInt(Int(Rnd() * 60))      ' Seconds 0-59
                                                                
                                Select Case mlngTimeFormat
                                       Case 23 To 32: intThousand = CInt(Int(Rnd() * 10000))  ' Thousands 0-9999  Simulate hi-perf timer
                                       Case Else:     intThousand = CInt(Int(Rnd() * 1000))   ' Thousands 0-999
                                End Select
                                
                                ' Date and time
                                If mlngDateType = 0 Then
                                    
                                    Select Case mlngDateFormat   ' Date format
                                           
                                           Case 0 To 15   ' Date and time with AM/PM
                                                lstRnd(0).AddItem Format$(Format$(DateSerial(intYear, intMonth, intDay), mstrDateFmt), strFmt1) & _
                                                                  " " & Format$(TimeSerial(intHour, intMinute, intSecond), strFmt2)
                                                  
                                           Case 16 To 27  ' Short Date and time without AM/PM
                                                lstRnd(0).AddItem Format$(Format$(DateSerial(intYear, intMonth, intDay), mstrDateFmt), strFmt1) & _
                                                                  " " & Format$(Format$(intHour, "00") & ":" & _
                                                                  Format$(intMinute, "00") & ":" & _
                                                                  Format$(intSecond, "00"), strFmt2)
                                           
                                           Case 28 To 32  ' Short Date and time with milliseconds and AM/PM
                                                lstRnd(0).AddItem Format$(Format$(DateSerial(intYear, intMonth, intDay), mstrDateFmt), strFmt1) & _
                                                                  " " & Format$(Format$(IIf(intHour > 12, intHour - 12, intHour), "00") & ":" & _
                                                                  Format$(intMinute, "00") & ":" & _
                                                                  Format$(intSecond, "00") & "." & _
                                                                  Format$(intThousand, strThou) & _
                                                                  IIf(intHour > 11, " PM", " AM"), strFmt2)
                                    
                                           Case 34 To 39  ' Short Date and time with milliseconds and AM/PM
                                                lstRnd(0).AddItem Format$(Format$(DateSerial(intYear, intMonth, intDay), mstrDateFmt), strFmt1) & _
                                                                  " " & Format$(Format$(IIf(intHour > 12, intHour - 12, intHour), "00") & ":" & _
                                                                  Format$(intMinute, "00") & ":" & _
                                                                  Format$(intSecond, "00") & "." & _
                                                                  Format$(intThousand, strThou), strFmt2)
                                    End Select
                                    
                                ' Date only
                                ElseIf mlngDateType = 1 Then
                                    
                                    lstRnd(0).AddItem Format$(Format$(DateSerial(intYear, intMonth, intDay), mstrDateFmt), strFmt1)

                                ' Time only
                                ElseIf mlngDateType = 2 Then
                                                                                            
                                    ' Time formatted to display in unsorted listbox
                                    Select Case mlngTimeFormat
                                           
                                           Case 0 To 5, 13 To 18
                                                lstRnd(0).AddItem Format$(Format$(TimeSerial(intHour, intMinute, intSecond), mstrTimeFmt), strFmt2)
                                                           
                                           Case 6, 19   ' With "A.M./P.M."
                                                lstRnd(0).AddItem Format$(Format$(IIf(intHour > 12, intHour - 12, intHour), "00") & strTemp & _
                                                                  Format$(intMinute, "00") & strTemp & _
                                                                  Format$(intSecond, "00") & _
                                                                  IIf(intHour > 11, " P.M.", " A.M."), strFmt2)
                                                                                      
                                           Case 7, 23
                                                lstRnd(0).AddItem Format$(Format$(IIf(intHour > 12, intHour - 12, intHour), "00") & strTemp & _
                                                                  Format$(intMinute, "00") & strTemp & _
                                                                  Format$(intSecond, "00") & "." & _
                                                                  Format$(intThousand, strThou), strFmt2)
                                                                  
                                           Case 8, 20, 24, 25
                                                lstRnd(0).AddItem Format$(Format$(IIf(intHour > 12, intHour - 12, intHour), "00") & strTemp & _
                                                                  Format$(intMinute, "00") & strTemp & _
                                                                  Format$(intSecond, "00") & strTemp & _
                                                                  Format$(intThousand, strThou), strFmt2)
                                                                  
                                           Case 9, 21, 26
                                                lstRnd(0).AddItem Format$(Format$(IIf(intHour > 12, intHour - 12, intHour), "00") & strTemp & _
                                                                  Format$(intMinute, "00") & strTemp & _
                                                                  Format$(intSecond, "00") & "." & _
                                                                  Format$(intThousand, strThou) & _
                                                                  IIf(intHour > 11, " PM", " AM"), strFmt2)
                                           
                                           Case 10, 27, 28
                                                lstRnd(0).AddItem Format$(Format$(IIf(intHour > 12, intHour - 12, intHour), "00") & strTemp & _
                                                                  Format$(intMinute, "00") & strTemp & _
                                                                  Format$(intSecond, "00") & strTemp & _
                                                                  Format$(intThousand, strThou) & _
                                                                  IIf(intHour > 11, " PM", " AM"), strFmt2)
                                           
                                           Case 11, 29
                                                lstRnd(0).AddItem Format$(Format$(IIf(intHour > 12, intHour - 12, intHour), "00") & strTemp & _
                                                                  Format$(intMinute, "00") & strTemp & _
                                                                  Format$(intSecond, "00") & "." & _
                                                                  Format$(intThousand, strThou) & _
                                                                  IIf(intHour > 11, " P.M.", " A.M."), strFmt2)
                                    
                                           Case 12, 22, 30, 31
                                                lstRnd(0).AddItem Format$(Format$(IIf(intHour > 12, intHour - 12, intHour), "00") & strTemp & _
                                                                  Format$(intMinute, "00") & strTemp & _
                                                                  Format$(intSecond, "00") & strTemp & _
                                                                  Format$(intThousand, strThou) & _
                                                                  IIf(intHour > 11, " P.M.", " A.M."), strFmt2)
                                    End Select
                                End If
                                
                            Next lngCount
                End Select
                                
                LockWindowUpdate 0&     ' Unlock listbox display after loading
                strRnd = vbNullString   ' Verify variable is empty (save memory resources)
                                
           Case 1  ' Sort listbox data
                ' make sure we have some data
                ' in the first listbox
                If lstRnd(0).ListCount < 1 Then
                    Exit Sub
                End If
                
                ' Deactivate so there are no more
                ' changes while data is being sorted
                fraDataType.Enabled = False
                cmdRnd(0).Enabled = False
                 
                Set objSort = New kiMisc.cSort   ' Instantiate sort class object
                lstRnd(1).Clear                  ' empty right side listbox
                
                LockWindowUpdate lstRnd(1).hwnd  ' Lock listbox to prevent flickering while loading
                    
                ' Determine data type selected
                Select Case mlngSelection
                       
                       Case 0  ' Numeric data
                            ReDim astrData(lstRnd(0).ListCount)  ' Size sorting array
                            
                            ' load numeric array
                            For lngIndex = 0 To lstRnd(0).ListCount - 1
                                astrData(lngIndex) = CStr(lstRnd(0).List(lngIndex))
                            Next lngIndex
                            
                            With objSort
                                .SortDirection = IIf(CBool(optSortOrder(0).Value), eSort_Ascending, eSort_Descending)
                                .TypeOfData = eSort_Numeric     ' Numeric data
                                .SortMethod = eQuickSort        ' using Quick Sort
                                .SortData astrData(), strTime   ' Sort data
                            End With
                                            
                            ' One way to determine how many duplicates were removed
                            If CBool(chkDuplicates.Value) Then
                                objSort.RemoveDupes astrData(), lngCount
                            End If
                            
                            ' Load second listbox with sorted data
                            For lngIndex = 0 To UBound(astrData) - 1
                                lstRnd(1).AddItem astrData(lngIndex)
                            Next lngIndex
                            
                       Case 1  ' String Data
                            ReDim astrData(lstRnd(0).ListCount)  ' Size sorting array
                            
                            ' load string array
                            For lngIndex = 0 To lstRnd(0).ListCount - 1
                                astrData(lngIndex) = CStr(lstRnd(0).List(lngIndex))
                            Next lngIndex
                            
                            With objSort
                                If CBool(chkCaseSensitive.Value) Then
                                    .CompareMethod = eSort_CaseSensitive
                                Else
                                    .CompareMethod = eSort_IgnoreCase
                                End If
                                
                                .SortDirection = IIf(CBool(optSortOrder(0).Value), eSort_Ascending, eSort_Descending)
                                .TypeOfData = eSort_String      ' String data
                                .SortMethod = eShellSort        ' using Shell Sort
                                .SortData astrData(), strTime   ' Sort data
                            End With
                            
                            ' remove duplicates
                            If CBool(chkDuplicates.Value) Then
                                objSort.RemoveDupes astrData()
                            End If
                            
                            ' Load second listbox with sorted data
                            For lngIndex = 0 To UBound(astrData) - 1
                                lstRnd(1).AddItem astrData(lngIndex)
                            Next lngIndex
                            
                            ' Another way to determine how many duplicates were removed
                            lngCount = lstRnd(0).ListCount - lstRnd(1).ListCount
                            
                       Case 2  ' Dates (Numeric sorting)
                            ReDim astrData(lstRnd(0).ListCount)   ' Size sorting array
                        
                            ' load string array
                            For lngIndex = 0 To lstRnd(0).ListCount - 1
                                astrData(lngIndex) = CStr(lstRnd(0).List(lngIndex))
                            Next lngIndex
                
                            With objSort
                                .SortDirection = IIf(CBool(optSortOrder(0).Value), eSort_Ascending, eSort_Descending)
                                .TypeOfData = eSort_Dates        ' Date sorting
                                .ProcessTime = IIf(lngDateTime = 2, False, True)
                                .DateFormat = mlngDateFormat     ' Date format
                                .TimeFormat = mlngTimeFormat     ' Time format
                                .SortMethod = eCombSort          ' using Comb Sort
                                .SortData astrData(), strTime    ' Sort data
                            End With

                            ' Determine how many duplicates were removed
                            If CBool(chkDuplicates.Value) Then
                                objSort.RemoveDupes astrData(), lngCount
                            End If
                            
                            ' Load second listbox with sorted data
                            For lngIndex = 0 To UBound(astrData) - 1
                                lstRnd(1).AddItem astrData(lngIndex)
                            Next lngIndex
                End Select
    
                LockWindowUpdate 0&    ' Unlock listbox display after loading
                lblRnd(2).Caption = "Elapsed time - " & strTime
    
                If CBool(chkDuplicates.Value) Then
                    lblRnd(5).Caption = "Sorted - " & CStr(lstRnd(1).ListCount) & Space$(5) & _
                                        CStr(lngCount) & " duplicates removed"
                Else
                    lblRnd(5).Caption = "Sorted - " & CStr(lstRnd(1).ListCount)
                End If
    End Select
    
    ' Always free objects from
    ' memory when not in use
    Set objRnd = Nothing
    Set objSort = Nothing
    
    ' Verify arrays are empty
    ' when not needed
    Erase abytData()
    Erase adblData()
    Erase astrData()
    
    ' Reset frames and command buttons
    fraSortOrder.Enabled = True
    fraDataType.Enabled = True
    cmdRnd(0).Enabled = True
    cmdRnd(1).Enabled = True

End Sub

Private Sub ShiftDemo(ByVal intIndex As Integer)

    ' Called by cmdBits_Click()
    
    Dim intIdx    As Integer
    Dim intBitPos As Integer
    Dim lngValue  As Long
    Dim lngTemp   As Long
    Dim objBits   As kiMisc.cMath32  ' Example of procedure level objects
    
    ' Empty the results boxes
    For intIdx = 1 To txtBitOutput.Count - 1
        txtBitOutput(intIdx).Text = vbNullString
    Next intIdx
        
    Set objBits = New kiMisc.cMath32  ' Instantiate object
    
    Select Case mlngValueType
           
           Case 0   ' Long Integer
                If Val(txtBitInput(0).Text) < MIN_LONG Or _
                   Val(txtBitInput(0).Text) > MAX_LONG Then
                   
                    InfoMsg "Long Integer data type has a range of" & vbNewLine & _
                            Format$(MIN_LONG, "##0") & " to " & Format$(MAX_LONG, "##0")
                    Exit Sub
                End If
    
           Case 1   ' Short Integer
                If Val(txtBitInput(0).Text) < MIN_INT Or _
                   Val(txtBitInput(0).Text) > MAX_INT Then
                   
                    InfoMsg "Short Integer data type has a range of" & vbNewLine & _
                            Format$(MIN_INT, "##0") & " to " & Format$(MAX_INT, "##0")
                    Exit Sub
                End If
    
           Case 2   ' Byte
                If Val(txtBitInput(0).Text) < 0 Or _
                   Val(txtBitInput(0).Text) > 255 Then
                   
                    InfoMsg "Byte data type has a range of 0 to 255."
                    Exit Sub
                End If
    End Select
    
    lngValue = Val(txtBitInput(0).Text)         ' Convert to numeric
    intBitPos = Abs(CInt(txtBitInput(1).Text))  ' Start with a positive value
    
    ' Test if number of shift positions exceed data type
    Select Case mlngValueType
        
           Case 0   ' Long Integer
                If intBitPos > 32 Then   ' Greater than long integer
                    InfoMsg "Number of shift positions exceed data type" & vbNewLine & vbNewLine & _
                            Space$(10) & "Long Integer - max 32"
                    Exit Sub
                End If
                
                If (lngValue < MIN_LONG) Or (lngValue > MAX_LONG) Then ' Exceeds value range
                    InfoMsg "Long Integer data type has a range of" & vbNewLine & _
                            Format$(MIN_LONG, "#,##0") & " to " & Format$(MAX_LONG, "#,##0")
                    Exit Sub
                End If
        
           Case 1   ' Short Integer
                If intBitPos > 16 Then   ' Greater than short integer
                    InfoMsg "Number of shift positions exceed data type" & vbNewLine & vbNewLine & _
                            Space$(10) & "Short Integer - max 16"
                    Exit Sub
                End If
            
                If (lngValue < MIN_INT) Or (lngValue > MAX_INT) Then  ' Exceeds value range
                    InfoMsg "Short Integer data type has a range of" & vbNewLine & _
                            Format$(MIN_INT, "##0") & " to " & Format$(MAX_INT, "##0")
                    Exit Sub
                End If
           
           Case 2   ' Byte
                If intBitPos > 8 Then   ' Greater than byte
                    InfoMsg "Number of shift positions exceed data type" & vbNewLine & vbNewLine & _
                            Space$(10) & "Byte - max 8"
                    Exit Sub
                End If
                
                If (lngValue < 0) Or (lngValue > 255) Then  ' Exceeds value range
                    InfoMsg "Byte data type has a positive range of 0-255"
                    Exit Sub
                End If
    End Select
    
    ' Shift or rotate data?
    Select Case intIndex
           
           Case 0   ' Shift data
                Select Case mlngValueType
                    
                       Case 0   ' Long Integer
                            ' Test if number of shift positions exceed data type
                            If intBitPos > 32 Then
                                
                                InfoMsg "Number of shift positions exceed data type" & vbNewLine & vbNewLine & _
                                        Space$(5) & Format$("Byte (max 8)", "!" & String$(25, "@")) & vbNewLine & _
                                        Space$(5) & Format$("Short Integer (max 16)", "!" & String$(25, "@")) & vbNewLine & _
                                        Space$(5) & Format$("Long Integer (max 32)", "!" & String$(25, "@"))
                                Exit Sub
                                
                            End If
        
                            ' Convert numeric to binary and adjust
                            ' the length of the return string
                            txtBitOutput(0).Text = objBits.NumberToBinary(lngValue)
                            txtBitOutput(0).Text = Right$(String$(32, "0") & txtBitOutput(0).Text, 32)
            
                            lngTemp = objBits.w32Shift(lngValue, intBitPos)          ' Shift left
                            txtBitOutput(1).Text = lngTemp                           ' display numeric value
                            txtBitOutput(2).Text = objBits.NumberToBinary(lngTemp)   ' Convert to binary string
                            txtBitOutput(2).Text = Right$(String$(32, "0") & txtBitOutput(2).Text, 32)
                            
                            intBitPos = -intBitPos                                   ' note negative bit position
                            lngTemp = objBits.w32Shift(lngValue, intBitPos)          ' Shift right
                            txtBitOutput(3).Text = lngTemp                           ' display numeric value
                            txtBitOutput(4).Text = objBits.NumberToBinary(lngTemp)   ' Convert to binary string
                            txtBitOutput(4).Text = Right$(String$(32, "0") & txtBitOutput(4).Text, 32)
                
                       Case 1   ' Short Integer
                            ' Test if number of shift positions exceed data type
                            If intBitPos > 16 Then
                                
                                InfoMsg "Number of shift positions exceed data type" & vbNewLine & vbNewLine & _
                                        Space$(5) & Format$("Byte (max 8)", "!" & String$(25, "@")) & vbNewLine & _
                                        Space$(5) & Format$("Short Integer (max 16)", "!" & String$(25, "@")) & vbNewLine & _
                                        Space$(5) & Format$("Long Integer (max 32)", "!" & String$(25, "@"))
                                Exit Sub
                                
                            End If
        
                            ' Convert numeric to binary and adjust
                            ' the length of the return string
                            txtBitOutput(0).Text = objBits.NumberToBinary(lngValue)
                            txtBitOutput(0).Text = Right$(String$(16, "0") & txtBitOutput(0).Text, 16)
            
                            lngTemp = objBits.w16Shift(lngValue, intBitPos)          ' Shift left
                            txtBitOutput(1).Text = lngTemp                           ' display numeric value
                            txtBitOutput(2).Text = objBits.NumberToBinary(lngTemp)   ' Convert to binary string
                            txtBitOutput(2).Text = Right$(String$(16, "0") & txtBitOutput(2).Text, 16)
                            
                            intBitPos = -intBitPos                                   ' note negative bit position
                            lngTemp = objBits.w16Shift(lngValue, intBitPos)          ' Shift right
                            txtBitOutput(3).Text = lngTemp                           ' display numeric value
                            txtBitOutput(4).Text = objBits.NumberToBinary(lngTemp)   ' Convert to binary string
                            txtBitOutput(4).Text = Right$(String$(16, "0") & txtBitOutput(4).Text, 16)
                    
                       Case 2   ' Byte
                            ' Convert numeric to binary and adjust
                            ' the length of the return string
                            txtBitOutput(0).Text = objBits.HexToBinary(Hex$(lngValue))
                            txtBitOutput(0).Text = Right$(String$(8, "0") & txtBitOutput(0).Text, 8)
        
                            lngTemp = objBits.w8Shift(CByte(lngValue), intBitPos)    ' Shift left
                            txtBitOutput(1).Text = lngTemp                           ' display numeric value
                            txtBitOutput(2).Text = objBits.NumberToBinary(lngTemp)   ' Convert to binary string
                            txtBitOutput(2).Text = Right$(String$(8, "0") & txtBitOutput(2).Text, 8)
                            
                            intBitPos = -intBitPos                                   ' note negative bit position
                            lngTemp = objBits.w8Shift(CByte(lngValue), intBitPos)    ' Shift right
                            txtBitOutput(3).Text = lngTemp                           ' display numeric value
                            txtBitOutput(4).Text = objBits.NumberToBinary(lngTemp)   ' Convert to binary string
                            txtBitOutput(4).Text = Right$(String$(8, "0") & txtBitOutput(4).Text, 8)
                End Select
                
           Case 1   ' Rotate data
                Select Case mlngValueType
                    
                       Case 0   ' Long Integer
                            ' Test if number of shift positions exceed data type
                            If intBitPos > 32 Then
                                
                                InfoMsg "Number of shift positions exceed data type" & vbNewLine & vbNewLine & _
                                        Space$(5) & Format$("Byte (max 8)", "!" & String$(25, "@")) & vbNewLine & _
                                        Space$(5) & Format$("Short Integer (max 16)", "!" & String$(25, "@")) & vbNewLine & _
                                        Space$(5) & Format$("Long Integer (max 32)", "!" & String$(25, "@"))
                                Exit Sub
                                
                            End If
        
                            ' Convert numeric to binary and adjust
                            ' the length of the return string
                            txtBitOutput(0).Text = objBits.HexToBinary(Hex$(lngValue))
                            txtBitOutput(0).Text = Right$(String$(32, "0") & txtBitOutput(0).Text, 32)
            
                            lngTemp = objBits.w32Rotate(lngValue, intBitPos)         ' Rotate left
                            txtBitOutput(1).Text = lngTemp                           ' display numeric value
                            txtBitOutput(2).Text = objBits.NumberToBinary(lngTemp)   ' Convert to binary string
                            txtBitOutput(2).Text = Right$(String$(32, "0") & txtBitOutput(2).Text, 32)
                            
                            intBitPos = -intBitPos                                   ' note negative bit position
                            lngTemp = objBits.w32Rotate(lngValue, intBitPos)         ' Rotate right
                            txtBitOutput(3).Text = lngTemp                           ' display numeric value
                            txtBitOutput(4).Text = objBits.NumberToBinary(lngTemp)   ' Convert to binary string
                            txtBitOutput(4).Text = Right$(String$(32, "0") & txtBitOutput(4).Text, 32)
                
                       Case 1   ' Short Integer
                            ' Test if number of shift positions exceed data type
                            If intBitPos > 16 Then
                                
                                InfoMsg "Number of shift positions exceed data type" & vbNewLine & vbNewLine & _
                                        Space$(5) & Format$("Byte (max 8)", "!" & String$(25, "@")) & vbNewLine & _
                                        Space$(5) & Format$("Short Integer (max 16)", "!" & String$(25, "@")) & vbNewLine & _
                                        Space$(5) & Format$("Long Integer (max 32)", "!" & String$(25, "@"))
                                Exit Sub
                                
                            End If
        
                            ' Convert numeric to binary and adjust
                            ' the length of the return string
                            txtBitOutput(0).Text = objBits.HexToBinary(Hex$(lngValue))
                            txtBitOutput(0).Text = Right$(String$(16, "0") & txtBitOutput(0).Text, 16)
            
                            lngTemp = objBits.w16Rotate(lngValue, intBitPos)         ' Rotate left
                            txtBitOutput(1).Text = lngTemp                           ' display numeric value
                            txtBitOutput(2).Text = objBits.NumberToBinary(lngTemp)   ' Convert to binary string
                            txtBitOutput(2).Text = Right$(String$(16, "0") & txtBitOutput(2).Text, 16)
                            
                            intBitPos = -intBitPos                                   ' note negative bit position
                            lngTemp = objBits.w16Rotate(lngValue, intBitPos)         ' Rotate right
                            txtBitOutput(3).Text = lngTemp                           ' display numeric value
                            txtBitOutput(4).Text = objBits.NumberToBinary(lngTemp)   ' Convert to binary string
                            txtBitOutput(4).Text = Right$(String$(16, "0") & txtBitOutput(4).Text, 16)
                       
                       Case 2   ' Byte
                            ' Convert numeric to binary and adjust
                            ' the length of the return string
                            txtBitOutput(0).Text = objBits.NumberToBinary(lngValue)
                            txtBitOutput(0).Text = Right$(String$(8, "0") & txtBitOutput(0).Text, 8)
            
                            lngTemp = objBits.w8Rotate(CByte(lngValue), intBitPos)   ' Rotate left
                            txtBitOutput(1).Text = lngTemp                           ' display numeric value
                            txtBitOutput(2).Text = objBits.NumberToBinary(lngTemp)   ' Convert to binary string
                            txtBitOutput(2).Text = Right$(String$(8, "0") & txtBitOutput(2).Text, 8)
                            
                            intBitPos = -intBitPos                                   ' note negative bit position
                            lngTemp = objBits.w8Rotate(CByte(lngValue), intBitPos)   ' Rotate right
                            txtBitOutput(3).Text = lngTemp                           ' display numeric value
                            txtBitOutput(4).Text = objBits.NumberToBinary(lngTemp)   ' Convert to binary string
                            txtBitOutput(4).Text = Right$(String$(8, "0") & txtBitOutput(4).Text, 8)
                End Select
    End Select
    
    Set objBits = Nothing   ' Always free objects from memory when not in use
    
End Sub

Private Sub DateDemo()

    ' Called by Form_Load()
    '           tmrDate_Timer()
    '           txtDate_LostFocus()
     
    Dim intIndex    As Integer
    Dim lngDay      As Long
    Dim lngMonth    As Long
    Dim lngCentury  As Long
    Dim datDate     As Date
    Dim strTemp     As String
    Dim strShortFmt As String
    Dim strLongFmt  As String
    
    ' Empty the results boxes
    For intIndex = 0 To txtDateResults.Count - 1
        txtDateResults(intIndex).Text = vbNullString
    Next intIndex
        
    lblDate(16).Caption = ""
    mstrPrevTime = Format$(Now, "hh:mm")
    lngMonth = 0
    lngDay = 0
    
    ' Load text boxes. Space preceeding or appending
    ' data prevents data from merging with frame
    ' while being displayed.
    With mobjDates
        strShortFmt = .ShortDateFormat  ' Get date formats for this locale
        strLongFmt = .LongDateFormat
        
        ' Date validated in
        ' txtDate_LostFocus() event
        datDate = CDate(txtDate.Text)   ' Convert to valid date variable

        ' Example of using VB FormatDateTime() function
        txtDateResults(0).Text = " " & Format$(datDate, strShortFmt)             ' Short date for this locale
        txtDateResults(1).Text = " " & Format$(datDate, strLongFmt)              ' Long date for this locale
        
        ' Example of using VB Format$() function
        txtDateResults(2).Text = " " & .DateToJulian(datDate)                    ' date to Julian
        lngCentury = CLng(Left$(CStr(Year(datDate)), 2) & "00")                  ' capture century
        strTemp = CStr(.JulianToDate(txtDateResults(2).Text, lngCentury))        ' Julian to date
        txtDateResults(3).Text = " " & Format$(CDate(strTemp), "dd MMM yyyy")    ' Modified format
        
        txtDateResults(4).Text = " " & Val(.DateToSerial(datDate))               ' date to serial
        strTemp = .SerialToDate(Val(txtDateResults(4).Text))                     ' serial to date
        txtDateResults(5).Text = " " & Format$(CDate(strTemp), "MMMM d, yyyy")   ' Long output
        
        txtDateResults(6).Text = .DateToWords(datDate)                           ' date to words
        
        ' Time display in 24 hour and 12 hour format
        txtDateResults(7).Text = " " & Format$(Now, "hh:mm:ss") & "  or  " & _
                                 Format$(Now, .TimeFormat)
        
        txtDateResults(8).Text = .TimeToWords(Format$(Now, "hh:mm:ss"))    ' Time to words
        
        ' Special day of month for year entered
        lblDate(12).Caption = "Sample dates for " & CStr(Year(datDate))
        Select Case cboDayOfMonth.ListIndex
        
               Case 0   ' Daylight Savings - Start
                    txtDateResults(11).Text = " " & .GetMonthName(3) & " " & _
                                              CStr(.DayOfMonth(2, 1, 3, Year(datDate)))
                    
                    If Len(.DaylightSavingsBegins(Year(datDate))) = 0 Then
                        txtDateResults(13).Text = " Observing standard time"
                        lblDate(16).Caption = "Daylight Savings Time"
                    Else
                        txtDateResults(13).Text = " " & .DaylightSavingsBegins(Year(datDate))
                        lblDate(16).Caption = "Daylight Savings Begins"
                    End If
                    
               Case 1   ' Daylight Savings - End
                    txtDateResults(11).Text = " " & .GetMonthName(11) & " " & _
                                              CStr(.DayOfMonth(1, 1, 11, Year(datDate)))
                    
                    If Len(.DaylightSavingsBegins(Year(datDate))) = 0 Then
                        txtDateResults(13).Text = " Observing standard time"
                        lblDate(16).Caption = "Daylight Savings Time"
                    Else
                        txtDateResults(13).Text = " " & .DaylightSavingsEnds(Year(datDate))
                        lblDate(16).Caption = "Daylight Savings Ends"
                    End If
                    
               Case 2   ' Martin Luther King Day
                    txtDateResults(11).Text = " " & .GetMonthName(1) & " " & _
                                               CStr(.DayOfMonth(3, 2, 1, Year(datDate)))
               Case 3   ' President's Day
                    txtDateResults(11).Text = " " & .GetMonthName(2) & " " & _
                                               CStr(.DayOfMonth(3, 2, 2, Year(datDate)))
               Case 4   ' Mother's Day
                    txtDateResults(11).Text = " " & .GetMonthName(5) & " " & _
                                              CStr(.DayOfMonth(2, 1, 5, Year(datDate)))
               Case 5   ' Father's Day
                    txtDateResults(11).Text = " " & .GetMonthName(6) & " " & _
                                              CStr(.DayOfMonth(3, 1, 6, Year(datDate)))
               Case 6   ' Thanksgiving Day (USA)
                    txtDateResults(11).Text = " " & .GetMonthName(11) & " " & _
                                              CStr(.DayOfMonth(4, 5, 11, Year(datDate)))
        End Select
        
        ' Get Easter Sunday for year entered
        datDate = .EasterSunday(Year(datDate))
        txtDateResults(12).Text = " " & .GetMonthName(Month(datDate)) & _
                                  " " & Format$(Day(datDate), "00")
        
        
        '****************************
        ' Y2K38 calculated values
        '****************************
        datDate = DateAdd("yyyy", 30, CDate(txtDate.Text))   ' Date +30 years
        
        If datDate > #1/18/2038# Then
            txtDateResults(9).Text = .Y2K38_DateToSeconds(datDate)   ' Date converted to seconds
            
            ' Seconds converted back to date
            txtDateResults(10).Text = Format$(.Y2K38_SecondsToDate(CDbl(txtDateResults(9).Text)), "dd-MMM-yyyy")
        End If
        
    End With
    
End Sub

Private Sub GetOperSystem()

    ' Called by Form_Load()
    
    Dim objOS As kiMisc.cOperSystem
    
    Set objOS = New kiMisc.cOperSystem   ' Instantiate object
    
    With objOS
        If .bWindowsNT Then
            lblOperSys.Caption = .VersionName & vbNewLine & _
                                 "Version " & .VersionNumber & _
                                 Space$(4) & .ServicePack
        Else
            lblOperSys.Caption = vbNullString
        End If
    End With
    
    Set objOS = Nothing   ' Always free objects when not needed
       
End Sub

Private Sub cboDate_Click()

    lstRnd(0).Clear   ' Clear listboxes
    lstRnd(1).Clear
    
    lblRnd(2).Caption = vbNullString
    lblRnd(5).Caption = vbNullString
    
    mlngDateType = cboDate.ListIndex
    
    Select Case mlngDateType
           Case 0, 1
                lblDateFmt.Caption = "Date format"   ' Date formats
                cboFormat(0).Visible = True
                cboFormat(1).Visible = False
           Case Else
                lblDateFmt.Caption = "Time format"   ' Time formats
                cboFormat(0).Visible = False
                cboFormat(1).Visible = True
    End Select
    
    cboFormat_Click (mlngDateType)

End Sub

Private Sub cboDayOfMonth_Click()
    DateDemo
End Sub

Private Sub cboDrive_Click()
    
    ' Called by LoadComboBoxes()
    
    Dim lngIndex    As Long
    Dim strDrive    As String
    Dim objDiskInfo As kiMisc.cDiskInfo
    
    strDrive = Trim$(cboDrive.Text)  ' Capture highlighted drive letter
    
    ' Clear data boxes
    For lngIndex = 0 To lblDrvData.Count - 1
        lblDrvData(lngIndex).Caption = vbNullString
    Next lngIndex
    
    ' Verify drive is available. This may
    ' be a flash drive that was removed.
    If Not IsPathValid(strDrive) Then
        InfoMsg strDrive & " is not available at this time." & Space$(5)
        LoadComboBoxes   ' Refresh drive letter combobox
        Exit Sub
    End If
    
    Set objDiskInfo = New kiMisc.cDiskInfo   ' Instantiate class object
    
    ' Load label boxes.  Space preceeding or appending
    ' data prevents data from merging with label frame
    ' while being displayed.
    With objDiskInfo
        .GetDriveInfo strDrive   ' Load properties with data
        
        ' Load property data onto form
        lblDrvData(0).Caption = " " & .DriveType
        lblDrvData(1).Caption = " " & .DriveTypeExtra
        lblDrvData(2).Caption = " " & .PartitionData
        lblDrvData(3).Caption = " " & .MfgModel
        lblDrvData(4).Caption = " " & .MfgSerial
        lblDrvData(5).Caption = " " & .MfgFirmware
        
        ' Disk information
        lblDrvData(6).Caption = Format$(.TotalDiskSpace, "#,##0") & " "
        lblDrvData(7).Caption = .DiskFormattedSize & " "
        lblDrvData(8).Caption = Format$(.DiskCylinders, "#,##0") & " "
        lblDrvData(9).Caption = Format$(.DiskTracksPerCyl, "#,##0") & " "
        lblDrvData(10).Caption = Format$(.DiskSectorsPerTrack, "#,##0") & " "
        lblDrvData(11).Caption = Format$(.DiskBytesPerSector, "#,##0") & " "
        
        ' Drive information (Partition)
        lblDrvData(12).Caption = " " & .DrvVolName
        lblDrvData(13).Caption = " " & .DrvVolSerial
        lblDrvData(14).Caption = " " & .DrvFileSystem
        lblDrvData(15).Caption = Format$(.DrvSectorsPerCluster, "#,##0") & " "
        lblDrvData(16).Caption = Format$(.DrvBytesPerSector, "#,##0") & " "
        lblDrvData(17).Caption = Format$(.DrvBytesPerCluster, "#,##0") & " "
        lblDrvData(18).Caption = .DrvFormattedSize & " "
        lblDrvData(19).Caption = Format$(.TotalDrvSpace, "#,##0") & " "
        lblDrvData(20).Caption = Format$(.DrvUsedSpace, "#,##0") & " "
        lblDrvData(21).Caption = Format$(.DrvAvailableSpace, "#,##0") & " "
        lblDrvData(22).Caption = Format$(.DrvFreeSpace, "#,##0") & " "
        lblDrvData(23).Caption = Format$(.DrvTotalClusters, "#,##0") & " "
        lblDrvData(24).Caption = Format$(.DrvFreeClusters, "#,##0") & " "
    End With
    
    Set objDiskInfo = Nothing   ' Always free object from memory when not needed
    
End Sub

Private Sub cboFormat_Click(Index As Integer)

    lstRnd(0).Clear   ' Clear listboxes
    lstRnd(1).Clear
        
    mstrDateFmt = vbNullString
    mstrTimeFmt = vbNullString
    mlngDateFormat = 0
    mlngTimeFormat = 0
    
    Select Case Index
            
           Case 0   ' Date format
                mlngDateFormat = cboFormat(0).ListIndex
                mstrDateFmt = cboFormat(0).Text
                
                If cboDate.ListIndex = 0 Then
                    
                    Select Case mlngDateFormat
                           Case 0 To 15
                                mlngTimeFormat = 3
                                mstrTimeFmt = String$(14, " ")   ' "hh:nn:ss AM/PM"
                           Case 16 To 27
                                mlngTimeFormat = 2
                                mstrTimeFmt = String$(8, " ")    ' "hh:nn:ss"
                           Case 28 To 33
                                mlngTimeFormat = 9
                                mstrTimeFmt = String$(18, " ")   ' "hh:nn:ss.ttt AM/PM"
                           Case 34 To 39
                                mlngTimeFormat = 23
                                mstrTimeFmt = String$(13, " ")   ' "hh:nn:ss.tttt"
                    End Select
                End If
                
           Case 1   ' Time only format
                mlngTimeFormat = cboFormat(1).ListIndex
                mstrTimeFmt = cboFormat(1).Text
    End Select

End Sub

Private Sub chkCaseSensitive_Click()
    
    lstRnd(1).Clear   ' Clear right side listbox
    
    lblRnd(2).Caption = vbNullString   ' Clear results
    lblRnd(5).Caption = vbNullString

End Sub

Private Sub chkDuplicates_Click()
    
    lstRnd(1).Clear   ' Clear right side listbox
    
    lblRnd(2).Caption = vbNullString   ' Clear results
    lblRnd(5).Caption = vbNullString

End Sub

Private Sub cmdBits_Click(Index As Integer)
    ShiftDemo Index
End Sub

' ***************************************************************************
' Routine:       cmdChoice_Click
'
' Description:   Performs the main functions of this application.  There is
'                STOP, GO, EXIT.
'
' Parameters:    Index - indicates which command button was clicked
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Private Sub cmdChoice_Click(Index As Integer)

    Select Case Index

           Case 0 ' about window
                frmMain.Hide
                frmAbout.DisplayAbout
           
           ' Shutdown this application
           Case Else
                gblnStopProcessing = True
                TerminateProgram
    End Select
    
End Sub

' ***************************************************************************
' Routine:       cmdConvert_Click
'
' Description:   Process whole numbers only
'
'                Maximum ranges for values that do not exceed a Long Integer.
'
'                               Numeric        Hex   Binary
'                Minimum:   -2147483648   80000000   10000000000000000000000000000000
'                Maximum:    2147483647   7FFFFFFF   01111111111111111111111111111111
'
'                Negative numeric values may be input.  Hex output is limited
'                to sixteen (16) characters.  Binary output is limited to
'                sixty-four (64) characters.  You may enter binary strings
'                that exceed sixty-four (64) characters but only the numeric
'                output value will be correct.  If a hex string entered converts
'                to a negative value, the numeric display will drop the minus
'                sign.  The binary display will be correct.
'
'                Below are the maximum ranges for all three displayed
'                values that exceed a Long Integer input value.
'
'                    Minimum:
'                Numeric   -9223372036854775808
'                Hex       8000000000000000  (if hex is entered, minus sign will be dropped from numeric output)
'                Binary    1000000000000000000000000000000000000000000000000000000000000000
'
'                    Maximum:
'                Numeric   9223372036854775807
'                Hex       7FFFFFFFFFFFFFFF
'                Binary    0111111111111111111111111111111111111111111111111111111111111111
'
' NOTE:          When converting negative values always double check your work.
'                Sometimes the conversions are accurate, other times not so.
'                Especially when dealing with hex.  Even the MS calculator can
'                produce erroneous results when performing conversions.
'
' MS Calc ex:     Number input:  -4794697086780616226    '<- Use negative number as input
'                   Hex output:  BD75D06728D751DE
'                Binary output:  1011110101110101110100000110011100101000110101110101000111011110
'
'                Number output:  13652046986928935390    ' Entirely different number is returned
'                    Hex input:  BD75D06728D751DE        '<- Now use hex value as input
'                Binary output:  1011110101110101110100000110011100101000110101110101000111011110
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub cmdConvert_Click(Index As Integer)

    Dim intIndex  As Integer
    Dim strBinary As String
    Dim objMath32 As kiMisc.cMath32
    Dim objMath64 As kiMisc.cMath64
    
    Set objMath32 = Nothing   ' Verify class objects are not active
    Set objMath64 = Nothing
    strBinary = vbNullString
    
    Select Case Index
            
           Case 0
                ' Remove leading and trailing blank spaces
                For intIndex = 0 To txtConvert.Count - 1
                    txtConvert(intIndex).Text = Trim$(txtConvert(intIndex).Text)
                Next intIndex
                
                ' Convert numeric data
                If Len(txtConvert(0).Text) > 0 Then
                
                    ' Check for decimals in numeric data
                    If InStr(1, txtConvert(0).Text, ".", vbBinaryCompare) > 0 Then
                        InfoMsg "Enter a valid whole number.  No decimals."
                        Exit Sub
                    End If
                    
                    ' Verify this is numeric data
                    If Not IsNumeric(txtConvert(0).Text) Then
                        InfoMsg "Enter numeric data only."
                        Exit Sub
                    End If
                        
                    ' See if numeric value exceeds a long integer
                    If Val(txtConvert(0).Text) > MAX_LONG Or _
                       Val(txtConvert(0).Text) < MIN_LONG Then
                        
                        Set objMath64 = New kiMisc.cMath64
                        With objMath64
                            txtConvert(1).Text = UCase$(.w64NumberToHex(txtConvert(0).Text))
                            txtConvert(2).Text = .w64NumberToBinary(txtConvert(0).Text)
                        End With
                       
                    Else
                        
                        ' Data is byte, short integer or long integer
                        Set objMath32 = New kiMisc.cMath32
                        With objMath32
                            txtConvert(1).Text = UCase$(.LongToHex(txtConvert(0).Text))
                            txtConvert(2).Text = .NumberToBinary(txtConvert(0).Text)
                        End With
                    End If
                        
                    GoTo cmdConvert_CleanUp
                
                End If
                
                ' Convert hex data
                If Len(txtConvert(1).Text) > 0 Then
                    
                    ' Remove unwanted items from hex data
                    txtConvert(1).Text = Replace(txtConvert(1).Text, "&", "")
                    txtConvert(1).Text = Replace(txtConvert(1).Text, "H", "")
                    
                    Set objMath64 = New kiMisc.cMath64  ' Instantiate class object
                    
                    ' Verify this is hex data
                    If Not objMath64.IsHexData(txtConvert(1).Text) Then
                        InfoMsg "This is not hex data."
                        txtConvert(1).SetFocus
                        GoTo cmdConvert_CleanUp
                    End If
                        
                    ' Are there more than eight hex characters?
                    If Len(txtConvert(1).Text) > 8 Then
                        With objMath64
                            txtConvert(0).Text = .w64HexToNumber(txtConvert(1).Text)
                            txtConvert(2).Text = .w64HexToBinary(txtConvert(1).Text)
                        End With
                    Else
                        ' Data is byte, short integer or long integer
                        Set objMath64 = Nothing      ' Deactivate class object
                        Set objMath32 = New kiMisc.cMath32  ' Instantiate class object
                    
                        With objMath32
                            txtConvert(0).Text = .HexToLong(txtConvert(1).Text)
                            txtConvert(2).Text = .HexToBinary(txtConvert(1).Text)
                        End With
                        
                    End If
                    
                    GoTo cmdConvert_CleanUp
                
                End If
                
                ' Convert binary data
                If Len(txtConvert(2).Text) > 0 Then
                    
                    Set objMath64 = New kiMisc.cMath64  ' Instantiate class object
                    
                    ' Clean up binary data string
                    strBinary = Trim$(txtConvert(2).Text)
                    strBinary = Replace(strBinary, Chr$(0), "")    ' Remove any null values
                    strBinary = Replace(strBinary, Chr$(32), "")   ' Remove any blank spaces
                    strBinary = Replace(strBinary, Chr$(13), "")   ' Remove any carriage returns
                    strBinary = Replace(strBinary, Chr$(10), "")   ' Remove any line feeds
                    
                    ' Verify this is binary data
                    If Not objMath64.IsBinaryData(strBinary) Then
                        InfoMsg "Binary data consist of zeroes and ones only."
                        txtConvert(2).SetFocus
                        GoTo cmdConvert_CleanUp
                    End If
                    
                    ' Does binary string length exceed thirty-two characters?
                    If Len(strBinary) > 32 Then
                        With objMath64
                            txtConvert(0).Text = .w64BinaryToNumber(strBinary)
                            txtConvert(1).Text = UCase$(.w64BinaryToHex(strBinary))
                        End With
                    Else
                        ' Data is byte, short integer or long integer
                        Set objMath64 = Nothing      ' Deactivate class object
                        Set objMath32 = New kiMisc.cMath32  ' Instantiate class object
                    
                        With objMath32
                            txtConvert(0).Text = .BinaryToNumber(strBinary)
                            txtConvert(1).Text = UCase$(.BinaryToHex(strBinary))
                        End With
                    End If
                    
                End If
           
           Case 1
                ' Empty conversion results boxes
                For intIndex = 0 To txtConvert.Count - 1
                    txtConvert(intIndex).Text = vbNullString
                Next intIndex
    End Select
    
cmdConvert_CleanUp:
    Set objMath32 = Nothing   ' Verify class objects are not active
    Set objMath64 = Nothing
    
End Sub

Private Sub cmdRefresh_Click()
    
    LoadComboBoxes   ' Refresh drive letter combobox

End Sub

Private Sub cmdRnd_Click(Index As Integer)

    Random_Demo Index

End Sub

Private Sub Form_Initialize()

    Set mobjDates = New kiMisc.cDates   ' instantiate date class (Example only)
    Set mobjKeyEdit = New cKeyEdit      ' Instantiate key edit class
       
End Sub

Private Sub Form_Load()

    Dim intIndex As Integer
    
    With frmMain
        .tmrDate.Enabled = False
        
        .Caption = PGM_NAME & gstrVersion
        mobjKeyEdit.CenterCaption frmMain
                
        optDataType_Click 0
        optValueType_Click 0
        
        .lblAuthor.Caption = AUTHOR_NAME
        .lblDate(15).Caption = "Date + 30 years must exceed 18 Jan 2038"
        .lblRnd(4).Caption = vbNullString
        .lblDateFmt.Caption = vbNullString
        
        .tabMain.TabIndex = 0
        .txtDate.Text = " " & Format$(Now(), "dd-Mmm-yyyy")  ' enter todays date
        
        .lstRnd(0).Clear
        .lstRnd(1).Clear
        
        ' Empty conversion results boxes
        For intIndex = 0 To .txtConvert.Count - 1
            .txtConvert(intIndex).Text = vbNullString
        Next intIndex

        LoadComboBoxes
        GetOperSystem
        
        ' center on the screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Hide
    End With
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    tmrDate.Enabled = False
    
    ' if kiMisc.DLL is still active
    ' then notify it to shut down
    If Not mobjDates Is Nothing Then
        mobjDates.StopProcessing = True
        DoEvents  ' allow it time to respond
    End If
    
    ' Always free objects from
    ' memory when not needed
    Set mobjDates = Nothing
    Set mobjKeyEdit = Nothing
    
    ' "X" in upper right corner was selected
    If UnloadMode = 0 Then
        TerminateProgram
    End If
    
End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub

Private Sub lblY2K38_Click()

    Dim strURL As String
    
    strURL = "http://visualstudiomagazine.com/articles/2010/02/16/the-next-y2k.aspx"
    
    ShellExecute GetDesktopWindow(), _
                 "open", strURL, 0&, 0&, SW_SHOWNORMAL
    
End Sub

Private Sub optDataType_Click(Index As Integer)
    
    lstRnd(0).Clear   ' Clear listboxes
    lstRnd(1).Clear
    
    lblRnd(2).Caption = vbNullString   ' Clear results
    lblRnd(5).Caption = vbNullString
    lblRnd(5).Caption = vbNullString
    mlngSelection = Index              ' Capture data type
    
    If chkCaseSensitive.Enabled Then
        chkCaseSensitive.Value = vbUnchecked
    End If
    
    Select Case Index
           Case 0   ' Numeric data
                optDataType(0).Value = True        ' Numeric data selected
                optDataType(1).Value = False
                optDataType(2).Value = False
                picDateFmt.Visible = False
                chkCaseSensitive.Enabled = False   ' Disable binary comparison
                lblRnd(3).Caption = "Sort method - Quick Sort"
                lblRnd(4).Caption = "Unsorted - " & CStr(SORT_ITEMS)

           Case 1   ' String data
                optDataType(0).Value = False
                optDataType(1).Value = True        ' String data selected
                optDataType(2).Value = False
                picDateFmt.Visible = False
                chkCaseSensitive.Enabled = True    ' Enable binary comparison
                lblRnd(3).Caption = "Sort method - Shell Sort"
                lblRnd(4).Caption = "Unsorted - " & CStr(SORT_ITEMS)
           
           Case 2   ' Dates
                optDataType(0).Value = False
                optDataType(1).Value = False
                optDataType(2).Value = True        ' Dates selected
                picDateFmt.Visible = True          ' Show date/time format boxes
                cboDate_Click                      ' Load combo boxes
                chkCaseSensitive.Enabled = False   ' Disable binary comparison
                lblRnd(3).Caption = "Sort method - Comb Sort"
                lblRnd(4).Caption = "Unsorted - " & CStr(SORT_ITEMS) & Space$(7) & "Dates: 1900-2199"

    End Select
    
End Sub

Private Sub optSortOrder_Click(Index As Integer)
    
    lstRnd(1).Clear                    ' Clear right side listbox
    lblRnd(2).Caption = vbNullString   ' Clear results
    lblRnd(5).Caption = vbNullString

End Sub

Private Sub optValueType_Click(Index As Integer)

    Dim intIndex As Integer
    
    mlngValueType = Index
    
    ' Empty result boxes
    For intIndex = 0 To txtBitOutput.Count - 1
        txtBitOutput(intIndex).Text = vbNullString
    Next intIndex
        
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)

    If tabMain.Tab = 0 Then
        tmrDate.Enabled = True
    Else
        tmrDate.Enabled = False
    End If

End Sub

Private Sub tmrDate_Timer()

    Dim strTime As String
    
    With mobjDates
        ' Time display in 24 hour and 12 hour format
        txtDateResults(7).Text = " " & Format$(Now, "hh:mm:ss") & "  or  " & _
                                 Format$(Now, .TimeFormat)
        
        strTime = Format$(Now, "hh:mm")
        
        ' See if current time has incremented to the next minute
        If StrComp(strTime, mstrPrevTime, vbTextCompare) = 1 Then
            mstrPrevTime = Format$(Now, "hh:mm")                          ' Save current time
            txtDateResults(8).Text = .TimeToWords(Format$(Now, "hh:mm"))  ' Time to words
        End If
    
        ' If the month does not equal the current month or the
        ' day no longer equals the current day then update all
        ' the text boxes with current data
        If Month(CDate(txtDate.Text)) <> Month(mdatDateEntered) Or _
           Day(CDate(txtDate.Text)) <> Day(mdatDateEntered) Then
            
            txtDate.Text = " " & Format$(CDate(txtDate.Text), "dd-MMM-yyyy")
            DateDemo
            Exit Sub
        End If
        
        ' If textbox is empty, use system date
        If Len(Trim$(txtDate.Text)) = 0 Then
            
            txtDate.Text = " " & Format$(Now(), "dd-MMM-yyyy")
            DateDemo
        End If
    End With
    
End Sub

' *************************************************************************
' Edit textbox data
' *************************************************************************
Private Sub txtBitInput_GotFocus(Index As Integer)

    Dim intIndex As Integer
    
    ' Empty the results boxes
    For intIndex = 0 To txtBitOutput.Count - 1
        txtBitOutput(intIndex).Text = vbNullString
    Next intIndex
        
    mobjKeyEdit.TextBoxFocus txtBitInput(Index)

End Sub

Private Sub txtBitInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    mobjKeyEdit.TextBoxKeyDown txtBitInput(Index), KeyCode, Shift
End Sub

Private Sub txtBitInput_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case KeyAscii
           Case 9                ' Tab key
                KeyAscii = 0
                SendKeys "{TAB}"
           Case 13               ' Enter key (no bell sound)
                KeyAscii = 0
           Case 8, 45, 48 To 57  ' backspace, minus, and numeric keys only
                ' good data
           Case Else             ' everything else
                KeyAscii = 0
    End Select
    
End Sub

Private Sub txtConvert_GotFocus(Index As Integer)
    mobjKeyEdit.TextBoxFocus txtConvert(Index)
End Sub

Private Sub txtConvert_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Dim strBinary As String
    
    mobjKeyEdit.TextBoxKeyDown txtConvert(Index), KeyCode, Shift

    ' if user is copying or cutting data
    ' from binary text box then verify
    ' data is formatted as binary
    If Index = 2 Then
    
        Select Case KeyCode
        
               ' C,X,c,x  Letter C or X pressed
               Case 67, 88, 99, 120
               
                    ' Control key also pressed
                    If Shift = 2 Then
                    
                        ' Clean up binary data string
                        strBinary = Clipboard.GetText()   ' unload clipboard into variable
                        
                        If Len(strBinary) > 0 Then
                            strBinary = Replace(strBinary, Chr$(0), "")    ' Remove any null values
                            strBinary = Replace(strBinary, Chr$(32), "")   ' Remove any blank spaces
                            strBinary = Replace(strBinary, Chr$(13), "")   ' Remove any carriage returns
                            strBinary = Replace(strBinary, Chr$(10), "")   ' Remove any line feeds
                                                                
                            Clipboard.Clear               ' clear clipboard contents
                            Clipboard.SetText strBinary   ' load clipboard with binary data
                        End If
                    End If
        End Select
        
    End If
    
End Sub

Private Sub txtConvert_KeyPress(Index As Integer, KeyAscii As Integer)
        
    Select Case Index
           Case 0   ' Numeric textbox
                Select Case KeyAscii
                       Case 9   ' Tab key
                            KeyAscii = 0
                            SendKeys "{TAB}"
                       Case 13: KeyAscii = 0     ' Enter key (no bell sound)
                       Case 8, 45, 48 To 57      ' backspace, minus, period, 0-9
                       Case Else: KeyAscii = 0   ' everything else
                End Select
                       
                ' Verify other textboxes are empty
                txtConvert(1).Text = vbNullString
                txtConvert(2).Text = vbNullString
                
           Case 1   ' Hex textbox
                Select Case KeyAscii
                       Case 9   ' Tab key
                            KeyAscii = 0
                            SendKeys "{TAB}"
                       Case 13: KeyAscii = 0                     ' Enter key (no bell sound)
                       Case 8, 48 To 57, 65 To 70                ' backspace, 0-9, A-F
                       Case 97 To 102: KeyAscii = KeyAscii - 32  ' a-f (Convert lowercase to uppercase)
                       Case Else: KeyAscii = 0                   ' everything else
                End Select
           
                ' Verify other textboxes are empty
                txtConvert(0).Text = vbNullString
                txtConvert(2).Text = vbNullString
                
           Case 2   ' Binary textbox
                Select Case KeyAscii
                       Case 9   ' Tab key
                            KeyAscii = 0
                            SendKeys "{TAB}"
                       Case 13: KeyAscii = 0     ' Enter key (no bell sound)
                       Case 8, 48, 49            ' backspace, 0, 1
                       Case Else: KeyAscii = 0   ' everything else
                End Select
    
                ' Verify other textboxes are empty
                txtConvert(0).Text = vbNullString
                txtConvert(1).Text = vbNullString
    End Select
    
End Sub

Private Sub txtConvert_LostFocus(Index As Integer)

    Select Case Index
            
           Case 0, 2  ' User updating numeric or binary textbox
                ' If there is something in text boxes
                If Len(txtConvert(Index).Text) > 0 Then
                    ' remove blank spaces (leading and trailing)
                    txtConvert(Index).Text = Trim$(txtConvert(Index).Text)
                End If
           
           Case 1   ' User updating hex textbox
                ' If there is something in the hex text box
                If Len(txtConvert(1).Text) > 0 Then
                    
                    ' remove blank spaces (leading and trailing)
                    txtConvert(1).Text = Trim$(txtConvert(1).Text)
                    
                    ' Pad leading zeroes
                    Select Case Len(txtConvert(1).Text)
                           Case Is <= 2:  txtConvert(1).Text = Right$("00" & txtConvert(1).Text, 2)               ' Byte
                           Case Is <= 4:  txtConvert(1).Text = Right$("0000" & txtConvert(1).Text, 4)             ' Short integer
                           Case Is <= 8:  txtConvert(1).Text = Right$(String$(8, "0") & txtConvert(1).Text, 8)    ' Long integer
                           Case Is <= 16: txtConvert(1).Text = Right$(String$(16, "0") & txtConvert(1).Text, 16)  ' Exceeds long integer
                    End Select
                End If
    End Select
    
End Sub

Private Sub txtDate_GotFocus()
    
    Dim intIndex As Integer
    
    tmrDate.Enabled = False  ' deactivate timer
    
    ' Empty the results boxes
    For intIndex = 0 To txtDateResults.Count - 1
        txtDateResults(intIndex).Text = vbNullString
    Next intIndex
        
    mobjKeyEdit.TextBoxFocus txtDate

End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    mobjKeyEdit.TextBoxKeyDown txtDate, KeyCode, Shift
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    mobjKeyEdit.ProcessAlphaNumeric KeyAscii
End Sub

Private Sub txtDate_LostFocus()

    Dim strTemp As String
    
    strTemp = Trim$(txtDate.Text)
    
    ' Empty text box
    If Len(strTemp) = 0 Then
    
        txtDate.Text = " " & FormatDateTime(Now(), vbShortDate)
        mdatDateEntered = CDate(txtDate.Text)   ' Save copy of date entered
        tmrDate.Enabled = True                  ' Activate timer
    End If
    
    ' Test for actual long date
    If IsDate(strTemp) Then
                 
        mdatDateEntered = CDate(txtDate.Text)   ' Save copy of date entered
        tmrDate.Enabled = True                  ' Activate timer
    End If
    
    ' Test for short date
    If IsDate("#" & strTemp & "#") Then
                 
        mdatDateEntered = CDate(txtDate.Text)   ' Save copy of date entered
        tmrDate.Enabled = True                  ' Activate timer
    End If
    
    ' User enter Now or Now()
    If StrComp(Left$(strTemp, 3), "now", vbTextCompare) = 0 Then
       
        txtDate.Text = " " & FormatDateTime(Now(), vbShortDate)
        mdatDateEntered = CDate(txtDate.Text)   ' Save copy of date entered
        tmrDate.Enabled = True                  ' Activate timer
    End If
    
    If tmrDate.Enabled Then
        DateDemo
    Else
        InfoMsg "Date must be in valid format for this locale."
        txtDate.SetFocus
        txtDate.Text = ""
    End If
    
End Sub

