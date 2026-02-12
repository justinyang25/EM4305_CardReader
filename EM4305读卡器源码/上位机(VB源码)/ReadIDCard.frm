VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "EM4305读写 https://shop72473760.taobao.com 石头工作室 13323889639"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ReadIDCard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6675
   ScaleWidth      =   9420
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame8 
      Caption         =   "写入块选项"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   6990
      TabIndex        =   125
      Top             =   1200
      Width           =   2385
      Begin VB.CheckBox Check3 
         Caption         =   "清空"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   1440
         TabIndex        =   139
         Top             =   930
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox Check3 
         Caption         =   "ALL"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   510
         TabIndex        =   138
         Top             =   930
         Width           =   765
      End
      Begin VB.CheckBox Check3 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   90
         TabIndex        =   137
         Top             =   930
         Width           =   555
      End
      Begin VB.CheckBox Check3 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   1920
         TabIndex        =   136
         Top             =   600
         Width           =   405
      End
      Begin VB.CheckBox Check3 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   1440
         TabIndex        =   135
         Top             =   600
         Width           =   555
      End
      Begin VB.CheckBox Check3 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   960
         TabIndex        =   134
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox Check3 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   510
         TabIndex        =   133
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox Check3 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   132
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox Check3 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   131
         Top             =   270
         Width           =   375
      End
      Begin VB.CheckBox Check3 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   130
         Top             =   270
         Width           =   375
      End
      Begin VB.CheckBox Check3 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   129
         Top             =   270
         Width           =   375
      End
      Begin VB.CheckBox Check3 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   510
         TabIndex        =   128
         Top             =   270
         Width           =   375
      End
      Begin VB.CheckBox Check3 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   127
         Top             =   270
         Width           =   375
      End
      Begin VB.CommandButton Command26 
         Caption         =   "写    卡"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   400
         TabIndex        =   126
         Top             =   1440
         Width           =   1680
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "配置位"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   90
      TabIndex        =   95
      Top             =   5040
      Width           =   6870
      Begin VB.CommandButton Command41 
         Caption         =   "缺省值"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4530
         TabIndex        =   142
         Top             =   1050
         Width           =   900
      End
      Begin VB.CheckBox Check2 
         Caption         =   "C26"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   6120
         TabIndex        =   124
         Top             =   270
         Width           =   585
      End
      Begin VB.CheckBox Check2 
         Caption         =   "C24"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   4905
         TabIndex        =   123
         Top             =   720
         Width           =   585
      End
      Begin VB.CheckBox Check2 
         Caption         =   "C23"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   4905
         TabIndex        =   122
         Top             =   495
         Width           =   585
      End
      Begin VB.CheckBox Check2 
         Caption         =   "C20"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   4905
         TabIndex        =   121
         Top             =   270
         Width           =   585
      End
      Begin VB.CheckBox Check2 
         Caption         =   "C18"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   3540
         TabIndex        =   120
         Top             =   1170
         Width           =   585
      End
      Begin VB.CommandButton Command27 
         Caption         =   "写配置字"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5550
         TabIndex        =   117
         Top             =   1050
         Width           =   1080
      End
      Begin VB.CheckBox Check2 
         Caption         =   "C17"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   3540
         TabIndex        =   116
         Top             =   945
         Width           =   585
      End
      Begin VB.CheckBox Check2 
         Caption         =   "C16"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   3540
         TabIndex        =   114
         Top             =   720
         Width           =   585
      End
      Begin VB.CheckBox Check2 
         Caption         =   "C15"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   3540
         TabIndex        =   112
         Top             =   510
         Width           =   585
      End
      Begin VB.CheckBox Check2 
         Caption         =   "C14"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3540
         TabIndex        =   110
         Top             =   270
         Width           =   585
      End
      Begin VB.CheckBox Check2 
         Caption         =   "BP/4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   2025
         TabIndex        =   108
         Top             =   1215
         Width           =   765
      End
      Begin VB.CheckBox Check2 
         Caption         =   "BP/8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   2025
         TabIndex        =   106
         Top             =   990
         Width           =   720
      End
      Begin VB.CheckBox Check2 
         Caption         =   "无延时"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   2025
         TabIndex        =   105
         Top             =   765
         Width           =   885
      End
      Begin VB.CheckBox Check2 
         Caption         =   "双相"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   2025
         TabIndex        =   104
         Top             =   510
         Width           =   840
      End
      Begin VB.CheckBox Check2 
         Caption         =   "曼切斯特"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2025
         TabIndex        =   103
         Top             =   270
         Width           =   1020
      End
      Begin VB.CheckBox Check2 
         Caption         =   "RF/64"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   630
         TabIndex        =   101
         Top             =   1230
         Width           =   825
      End
      Begin VB.CheckBox Check2 
         Caption         =   "RF/40"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   630
         TabIndex        =   100
         Top             =   990
         Width           =   765
      End
      Begin VB.CheckBox Check2 
         Caption         =   "RF/32"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   630
         TabIndex        =   99
         Top             =   750
         Width           =   825
      End
      Begin VB.CheckBox Check2 
         Caption         =   "RF/16"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   630
         TabIndex        =   98
         Top             =   510
         Width           =   765
      End
      Begin VB.CheckBox Check2 
         Caption         =   "RF/8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   630
         TabIndex        =   97
         Top             =   270
         Width           =   675
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "延时："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   7
         Left            =   1485
         TabIndex        =   119
         Top             =   810
         Width           =   480
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "编码："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   6
         Left            =   1485
         TabIndex        =   118
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Mode："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   5625
         TabIndex        =   115
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "RTF："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   4500
         TabIndex        =   113
         Top             =   765
         Width           =   450
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "禁  用："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   4230
         TabIndex        =   111
         Top             =   540
         Width           =   720
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "写登录："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   4230
         TabIndex        =   109
         Top             =   315
         Width           =   720
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "读登录："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   2880
         TabIndex        =   107
         Top             =   1215
         Width           =   720
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "LWR："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   3150
         TabIndex        =   102
         Top             =   315
         Width           =   450
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "速率："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   96
         Top             =   300
         Width           =   540
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4140
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      InputMode       =   1
   End
   Begin VB.Frame Frame6 
      Caption         =   "ID卡复制"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   3840
      TabIndex        =   71
      Top             =   60
      Width           =   5550
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   1140
         Top             =   810
      End
      Begin VB.CommandButton Command40 
         Caption         =   "读  卡"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2565
         TabIndex        =   78
         Top             =   720
         Width           =   1230
      End
      Begin VB.CommandButton Command39 
         Caption         =   "写  卡"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4185
         TabIndex        =   77
         Top             =   720
         Width           =   1275
      End
      Begin VB.CheckBox Check10 
         Caption         =   "递增"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4770
         TabIndex        =   76
         Top             =   330
         Value           =   1  'Checked
         Width           =   825
      End
      Begin VB.TextBox Text18 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3450
         TabIndex        =   75
         Text            =   "0000000000"
         Top             =   315
         Width           =   1290
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "ReadIDCard.frx":0CCE
         Left            =   1125
         List            =   "ReadIDCard.frx":0CD8
         TabIndex        =   72
         Text            =   "T5557"
         Top             =   315
         Width           =   990
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "16进制卡号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2295
         TabIndex        =   74
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "基卡类型："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   73
         Top             =   360
         Width           =   1050
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "EM4100卡读"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   1830
      TabIndex        =   67
      Top             =   60
      Width           =   2010
      Begin VB.CommandButton Command38 
         Caption         =   "读 卡"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   675
         TabIndex        =   70
         Top             =   720
         Width           =   1140
      End
      Begin VB.TextBox Text17 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   675
         TabIndex        =   69
         Text            =   "0000000000"
         Top             =   360
         Width           =   1290
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   240
         Shape           =   3  'Circle
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "卡号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   68
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "操作"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   6990
      TabIndex        =   63
      Top             =   3510
      Width           =   2385
      Begin VB.CommandButton Command37 
         Caption         =   "初始化卡"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   400
         TabIndex        =   141
         Top             =   2390
         Width           =   1680
      End
      Begin VB.CommandButton Command36 
         Caption         =   "写保护位"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   400
         TabIndex        =   66
         Top             =   1730
         Width           =   1680
      End
      Begin VB.CommandButton Command35 
         Caption         =   "休眠卡片"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   400
         TabIndex        =   65
         Top             =   1080
         Width           =   1680
      End
      Begin VB.CommandButton Command33 
         Caption         =   "读    卡"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   400
         TabIndex        =   64
         Top             =   420
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "EM4305/4205卡EEprom块"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3810
      Left            =   90
      TabIndex        =   1
      Top             =   1200
      Width           =   6870
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "ReadIDCard.frx":0CEB
         Left            =   1320
         List            =   "ReadIDCard.frx":0CFE
         TabIndex        =   145
         Top             =   3330
         Width           =   1170
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3915
         TabIndex        =   143
         Text            =   "00000000"
         Top             =   3360
         Width           =   930
      End
      Begin VB.CommandButton Command34 
         Caption         =   "修   改"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5445
         TabIndex        =   140
         Top             =   3360
         Width           =   1200
      End
      Begin VB.CheckBox Check1 
         Caption         =   "锁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   4920
         TabIndex        =   94
         Top             =   1860
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "锁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   4920
         TabIndex        =   93
         Top             =   1500
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "锁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   4920
         TabIndex        =   92
         Top             =   1140
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "锁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   4920
         TabIndex        =   91
         Top             =   780
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "锁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   4920
         TabIndex        =   90
         Top             =   420
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "锁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   1500
         TabIndex        =   89
         Top             =   2940
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "锁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   1500
         TabIndex        =   88
         Top             =   2580
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "锁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   1500
         TabIndex        =   87
         Top             =   2220
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "锁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   1500
         TabIndex        =   86
         Top             =   1860
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "锁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   1500
         TabIndex        =   85
         Top             =   1500
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "锁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1500
         TabIndex        =   84
         Top             =   1170
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "锁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   83
         Top             =   2580
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "锁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1500
         TabIndex        =   82
         Top             =   810
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "锁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   81
         Top             =   2970
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "锁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   80
         Top             =   2220
         Width           =   480
      End
      Begin VB.CheckBox Check1 
         Caption         =   "锁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1500
         TabIndex        =   79
         Top             =   450
         Width           =   465
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3915
         TabIndex        =   52
         Text            =   "00000000"
         Top             =   2970
         Width           =   930
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         HideSelection   =   0   'False
         Left            =   3915
         TabIndex        =   51
         Text            =   "00000000"
         Top             =   2565
         Width           =   930
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3915
         TabIndex        =   50
         Text            =   "00000000"
         Top             =   2205
         Width           =   930
      End
      Begin VB.TextBox Text13 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3915
         TabIndex        =   49
         Text            =   "00000000"
         Top             =   1845
         Width           =   930
      End
      Begin VB.TextBox Text12 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3915
         TabIndex        =   48
         Text            =   "00000000"
         Top             =   1485
         Width           =   930
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3915
         TabIndex        =   47
         Text            =   "00000000"
         Top             =   1125
         Width           =   930
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3915
         TabIndex        =   46
         Text            =   "00000000"
         Top             =   765
         Width           =   930
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3915
         TabIndex        =   45
         Text            =   "00000000"
         Top             =   405
         Width           =   930
      End
      Begin VB.CommandButton Command32 
         Caption         =   "写"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5445
         TabIndex        =   44
         Top             =   375
         Width           =   555
      End
      Begin VB.CommandButton Command31 
         Caption         =   "写"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5445
         TabIndex        =   43
         Top             =   750
         Width           =   555
      End
      Begin VB.CommandButton Command30 
         Caption         =   "写"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5445
         TabIndex        =   42
         Top             =   1140
         Width           =   555
      End
      Begin VB.CommandButton Command25 
         Caption         =   "登   录"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5445
         TabIndex        =   41
         Top             =   2940
         Width           =   1200
      End
      Begin VB.CommandButton Command24 
         Caption         =   "读"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6075
         TabIndex        =   40
         Top             =   375
         Width           =   555
      End
      Begin VB.CommandButton Command23 
         Caption         =   "读"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6075
         TabIndex        =   39
         Top             =   750
         Width           =   555
      End
      Begin VB.CommandButton Command22 
         Caption         =   "读"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6075
         TabIndex        =   38
         Top             =   1125
         Width           =   555
      End
      Begin VB.CommandButton Command21 
         Caption         =   "读"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5445
         TabIndex        =   37
         Top             =   1500
         Width           =   1185
      End
      Begin VB.CommandButton Command20 
         Caption         =   "读"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5445
         TabIndex        =   36
         Top             =   1860
         Width           =   1185
      End
      Begin VB.CommandButton Command19 
         Caption         =   "读卡号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5445
         TabIndex        =   35
         Top             =   2220
         Width           =   1185
      End
      Begin VB.CommandButton Command18 
         Caption         =   " 读配置 "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5445
         TabIndex        =   34
         Top             =   2580
         Width           =   1200
      End
      Begin VB.CommandButton Command16 
         Caption         =   "读"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2655
         TabIndex        =   33
         Top             =   2925
         Width           =   555
      End
      Begin VB.CommandButton Command15 
         Caption         =   "读"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2655
         TabIndex        =   32
         Top             =   2565
         Width           =   555
      End
      Begin VB.CommandButton Command14 
         Caption         =   "读"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2655
         TabIndex        =   31
         Top             =   2205
         Width           =   555
      End
      Begin VB.CommandButton Command13 
         Caption         =   "读"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2655
         TabIndex        =   30
         Top             =   1845
         Width           =   555
      End
      Begin VB.CommandButton Command12 
         Caption         =   "读"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2655
         TabIndex        =   29
         Top             =   1485
         Width           =   555
      End
      Begin VB.CommandButton Command11 
         Caption         =   "读"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2655
         TabIndex        =   28
         Top             =   1125
         Width           =   555
      End
      Begin VB.CommandButton Command10 
         Caption         =   "读"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2655
         TabIndex        =   27
         Top             =   765
         Width           =   555
      End
      Begin VB.CommandButton Command9 
         Caption         =   "读"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2655
         TabIndex        =   26
         Top             =   405
         Width           =   555
      End
      Begin VB.CommandButton Command8 
         Caption         =   "写"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2025
         TabIndex        =   25
         Top             =   2925
         Width           =   555
      End
      Begin VB.CommandButton Command7 
         Caption         =   "写"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2025
         TabIndex        =   24
         Top             =   2565
         Width           =   555
      End
      Begin VB.CommandButton Command6 
         Caption         =   "写"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2025
         TabIndex        =   23
         Top             =   2205
         Width           =   555
      End
      Begin VB.CommandButton Command5 
         Caption         =   "写"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2025
         TabIndex        =   22
         Top             =   1845
         Width           =   555
      End
      Begin VB.CommandButton Command4 
         Caption         =   "写"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2025
         TabIndex        =   21
         Top             =   1485
         Width           =   555
      End
      Begin VB.CommandButton Command3 
         Caption         =   "写"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2025
         TabIndex        =   20
         Top             =   1125
         Width           =   555
      End
      Begin VB.CommandButton Command2 
         Caption         =   "写"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2025
         TabIndex        =   19
         Top             =   765
         Width           =   555
      End
      Begin VB.CommandButton Command1 
         Caption         =   "写"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2025
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   405
         Width           =   555
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   495
         TabIndex        =   9
         Text            =   "00000000"
         Top             =   405
         Width           =   930
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   8
         Text            =   "00000000"
         Top             =   765
         Width           =   930
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   7
         Text            =   "00000000"
         Top             =   1125
         Width           =   930
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   6
         Text            =   "00000000"
         Top             =   1500
         Width           =   930
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   5
         Text            =   "00000000"
         Top             =   1845
         Width           =   930
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   4
         Text            =   "00000000"
         Top             =   2205
         Width           =   930
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   3
         Text            =   "00000000"
         Top             =   2565
         Width           =   930
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   495
         TabIndex        =   2
         Text            =   "00000000"
         Top             =   2925
         Width           =   930
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "基卡速率："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   146
         Top             =   3390
         Width           =   1275
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "新密码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3210
         TabIndex        =   144
         Top             =   3390
         Width           =   675
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "密码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3420
         TabIndex        =   60
         Top             =   3000
         Width           =   450
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "配置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3420
         TabIndex        =   59
         Top             =   2565
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "UID"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3465
         TabIndex        =   58
         Top             =   2250
         Width           =   315
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "块15"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3420
         TabIndex        =   57
         Top             =   1890
         Width           =   420
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "块14"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3420
         TabIndex        =   56
         Top             =   1530
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "块13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3420
         TabIndex        =   55
         Top             =   1170
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "块12"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3420
         TabIndex        =   54
         Top             =   810
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "块11"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3420
         TabIndex        =   53
         Top             =   450
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "块0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   17
         Top             =   450
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "块3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   16
         Top             =   795
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "块5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   15
         Top             =   1170
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "块6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   14
         Top             =   1515
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "块7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   13
         Top             =   1875
         Width           =   315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "块8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   12
         Top             =   2235
         Width           =   315
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "块9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   11
         Top             =   2595
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "块10"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   10
         Top             =   2955
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "串口设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   1740
      Begin VB.CommandButton Command17 
         Caption         =   "打 开"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   720
         TabIndex        =   62
         Top             =   720
         Width           =   870
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "ReadIDCard.frx":0D24
         Left            =   135
         List            =   "ReadIDCard.frx":0D26
         TabIndex        =   61
         Top             =   315
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   180
         Shape           =   3  'Circle
         Top             =   765
         Width           =   300
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UserBuff(1024) As Byte
Dim UserRxdBuff(1024) As Byte
Dim ReturnData(1) As Byte
Dim WriteDel As Byte
Dim PassTimeS As Byte
Private Declare Sub Sleep Lib "kernel32.DLL" (ByVal dwMilliseconds As Long)


Public Function IsHex(c As String) As Integer
If c >= "0" And c <= "9" Then
  IsHex = Val(c) - Val("0")
ElseIf c >= "a" And c <= "f" Then
  IsHex = Asc(c) - Asc("a") + 10
ElseIf c >= "A" And c <= "F" Then
  IsHex = Asc(c) - Asc("A") + 10
Else
  IsHex = 16
End If
End Function

Function CRC16(data() As Byte, L As Integer) As String
    Dim CRC16Lo As Byte, CRC16Hi As Byte      'CRC寄存器
    Dim CL As Byte, CH As Byte                '多项式码&HA001
    Dim SaveHi As Byte, SaveLo As Byte
    Dim i As Integer
    Dim Flag As Integer
    
    CRC16Lo = &HFF
    CRC16Hi = &HFF
    
    CL = &H1
    CH = &HA0
    
    For i = 0 To L
        CRC16Lo = CRC16Lo Xor data(i)        '每一个数据与CRC寄存器进行异或
        For Flag = 0 To 7
            SaveHi = CRC16Hi
            SaveLo = CRC16Lo
            CRC16Hi = CRC16Hi \ 2            '高位右移一位
            CRC16Lo = CRC16Lo \ 2            '低位右移一位
            If ((SaveHi And &H1) = &H1) Then '如果高位字节最后一位为1
                CRC16Lo = CRC16Lo Or &H80      '则低位字节右移后前面补1
            End If                           '否则自动补0
            If ((SaveLo And &H1) = &H1) Then '如果LSB为1，则与多项式码进行异或
                CRC16Hi = CRC16Hi Xor CH
                CRC16Lo = CRC16Lo Xor CL
            End If
        Next Flag
    Next i
    ReturnData(0) = CRC16Hi              'CRC高位
    ReturnData(1) = CRC16Lo              'CRC低位
End Function

Private Sub EM4100_Updata()     '配置块显示

    Dim cnt As Integer
    Dim s As String
    Dim L As Integer
    Dim Up_data(5) As Byte

    s = Trim(Text18.Text)
    L = Len(s)
    cnt = 0
    Do While (L)
        Do Until IsHex(Left(s, 1)) <> 16
            L = L - 1
            If L = 0 Then Exit Do
                s = Right(s, L)
        Loop
        a = IsHex(Left(s, 1))
        L = L - 1
        s = Right(s, L)
        If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
            a = a * 16 + IsHex(Left(s, 1))
            L = L - 1
            s = Right(s, L)
        End If
        
        Up_data(cnt) = a
        cnt = cnt + 1
    Loop
    If Up_data(4) = &HFF Then
        Up_data(4) = 0
        If Up_data(3) = &HFF Then
            Up_data(3) = 0
                If Up_data(2) = &HFF Then
                    Up_data(2) = 0
                    If Up_data(1) = &HFF Then
                        If Up_data(0) = &HFF Then
                                Up_data(0) = 0
                                Up_data(1) = 0
                                Up_data(2) = 0
                                Up_data(3) = 0
                                Up_data(4) = 1
                        Else
                            Up_data(0) = Up_data(0) + 1
                        End If
                    Else
                        Up_data(1) = Up_data(1) + 1
                    End If
                Else
                    Up_data(2) = Up_data(2) + 1
                End If
        Else
            Up_data(3) = Up_data(3) + 1
        End If
    Else
        Up_data(4) = Up_data(4) + 1
    End If
    
    Text18.Text = ""
    Text18.Text = Right("0" + Hex(Up_data(0)), 2) '十六进制收数
    Text18.Text = Text18.Text & Right("0" + Hex(Up_data(1)), 2)  '十六进制收数
    Text18.Text = Text18.Text & Right("0" + Hex(Up_data(2)), 2)  '十六进制收数
    Text18.Text = Text18.Text & Right("0" + Hex(Up_data(3)), 2)  '十六进制收数
    Text18.Text = Text18.Text & Right("0" + Hex(Up_data(4)), 2)  '十六进制收数
End Sub

Private Sub Bank4Data()     '配置块显示
    Dim Bank4_data(4) As Byte
    Bank4_data(0) = 0
    Bank4_data(1) = 0
    Bank4_data(2) = 0
    Bank4_data(3) = 0
    
    If Check2(0) = 1 Then
         Bank4_data(0) = &H3
    ElseIf Check2(1) = 1 Then
         Bank4_data(0) = &H7
    ElseIf Check2(2) = 1 Then
         Bank4_data(0) = &HF
    ElseIf Check2(3) = 1 Then
         Bank4_data(0) = &H13
    ElseIf Check2(4) = 1 Then
         Bank4_data(0) = &H1F
    End If
    
    If Check2(5) = 1 Then      '编码
        Bank4_data(0) = Bank4_data(0) Or &H40
    ElseIf Check2(1) = 1 Then
        Bank4_data(0) = Bank4_data(0) Or &H80
    End If
    
    If Check2(7) = 1 Then     '延时
        Bank4_data(1) = &H0
    ElseIf Check2(8) = 1 Then
        Bank4_data(1) = &H20
    ElseIf Check2(9) = 1 Then
        Bank4_data(1) = &H10
    End If

    If Check2(10) = 1 Then          'LWR
        Bank4_data(1) = Bank4_data(1) Or &H40
    End If
    If Check2(11) = 1 Then
        Bank4_data(1) = Bank4_data(1) Or &H80
    End If
    If Check2(12) = 1 Then
        Bank4_data(2) = Bank4_data(2) Or &H1
    End If
    If Check2(13) = 1 Then
        Bank4_data(2) = Bank4_data(2) Or &H2
    End If
    
    If Check2(14) = 1 Then          '读登陆
        Bank4_data(2) = Bank4_data(2) Or &H4
    End If
    If Check2(15) = 1 Then          '写登陆
        Bank4_data(2) = Bank4_data(2) Or &H10
    End If
    If Check2(16) = 1 Then          '禁用
        Bank4_data(2) = Bank4_data(2) Or &H80
    End If
    If Check2(17) = 1 Then          'RTF
        Bank4_data(3) = Bank4_data(3) Or &H1
    End If
    If Check2(18) = 1 Then           'Mode
        Bank4_data(3) = Bank4_data(3) Or &H4
    End If
    
    Text15.Text = Right("0" + Hex(Bank4_data(3)), 2) '十六进制收数
    Text15.Text = Text15.Text & Right("0" + Hex(Bank4_data(2)), 2)  '十六进制收数
    Text15.Text = Text15.Text & Right("0" + Hex(Bank4_data(1)), 2)  '十六进制收数
    Text15.Text = Text15.Text & Right("0" + Hex(Bank4_data(0)), 2)  '十六进制收数
End Sub

Private Sub UartBank4Data()     '配置块显示
    Dim Bank4_data As Byte   '4 5 6 7
     
    Bank4_data = UserRxdBuff(7) And &H1F
    Check2(0).Value = 0
    Check2(1).Value = 0
    Check2(2).Value = 0
    Check2(3).Value = 0
    Check2(4).Value = 0
    
    Select Case Bank4_data                '速率
            Case &H3
                    Check2(0).Value = 1
            Case &H7
                    Check2(1).Value = 1
            Case &HF
                    Check2(2).Value = 1
            Case &H13
                    Check2(3).Value = 1
            Case &H1F
                    Check2(4).Value = 1
    End Select
    
    Bank4_data = UserRxdBuff(7) And &HC0
    
    Select Case Bank4_data                '编码方式
            Case &H40
                    Check2(5).Value = 1
                    Check2(6).Value = 0
            Case &H80
                    Check2(6).Value = 1
                    Check2(5).Value = 0
    End Select
    
    Bank4_data = UserRxdBuff(6) And &H30
    Select Case Bank4_data                '延时
            Case &H0
                    Check2(7).Value = 1
                    Check2(8).Value = 0
                    Check2(9).Value = 0
            Case &H30
                    Check2(7).Value = 1
                    Check2(8).Value = 0
                    Check2(9).Value = 0
            Case &H20
                    Check2(8).Value = 1
                    Check2(7).Value = 0
                    Check2(9).Value = 0
            Case &H10
                    Check2(9).Value = 1
                    Check2(7).Value = 0
                    Check2(8).Value = 0
    End Select
    
    Bank4_data = UserRxdBuff(6) And &HC0
    
    Select Case Bank4_data                'LWR
            Case &HC0
                    Check2(10).Value = 1
                    Check2(11).Value = 1
            Case &H40
                    Check2(10).Value = 1
                    Check2(11).Value = 0
            Case &H80
                    Check2(10).Value = 0
                    Check2(11).Value = 1
    End Select
    
    Bank4_data = UserRxdBuff(5) And &H3
    Select Case Bank4_data                'LWR
            Case &H3
                    Check2(12).Value = 1
                    Check2(13).Value = 1
            Case &H1
                    Check2(12).Value = 1
                    Check2(13).Value = 0
            Case &H2
                    Check2(12).Value = 0
                    Check2(13).Value = 1
    End Select
    
    Bank4_data = UserRxdBuff(5) And &H4   '读登录
    If Bank4_data = &H4 Then
        Check2(14).Value = 1
    Else
        Check2(14).Value = 0
    End If
    
    Bank4_data = UserRxdBuff(5) And &H10  '写登录
    If Bank4_data = &H10 Then
        Check2(15).Value = 1
    Else
        Check2(15).Value = 0
    End If
     
    Bank4_data = UserRxdBuff(5) And &H80  '禁用
    If Bank4_data = &H80 Then
        Check2(16).Value = 1
    Else
        Check2(16).Value = 0
    End If
    
    Bank4_data = UserRxdBuff(4) And &H1  'RTF
    If Bank4_data = &H1 Then
        Check2(17).Value = 1
    Else
        Check2(17).Value = 0
    End If
    
    Bank4_data = UserRxdBuff(4) And &H4  'MODE
    If Bank4_data = &H4 Then
        Check2(18).Value = 1
    Else
        Check2(18).Value = 0
    End If
End Sub
Private Sub UartBank14Data()

    If (UserRxdBuff(7) And &H1) = &H1 Then
        Check1(0).Value = 1
    Else
        Check1(0).Value = 0
    End If
    
    If (UserRxdBuff(7) And &H2) = &H2 Then
        Check1(1).Value = 1
    Else
        Check1(1).Value = 0
    End If
    
    If (UserRxdBuff(7) And &H4) = &H4 Then
        Check1(2).Value = 1
    Else
        Check1(2).Value = 0
    End If
    
    If (UserRxdBuff(7) And &H8) = &H8 Then
        Check1(3).Value = 1
    Else
        Check1(3).Value = 0
    End If
    
    If (UserRxdBuff(7) And &H10) = &H10 Then
        Check1(4).Value = 1
    Else
        Check1(4).Value = 0
    End If
    If (UserRxdBuff(7) And &H20) = &H20 Then
        Check1(5).Value = 1
    Else
        Check1(5).Value = 0
    End If
    If (UserRxdBuff(7) And &H40) = &H40 Then
        Check1(6).Value = 1
    Else
        Check1(6).Value = 0
    End If
    If (UserRxdBuff(7) And &H80) = &H80 Then
        Check1(7).Value = 1
    Else
        Check1(7).Value = 0
    End If
    
    If (UserRxdBuff(6) And &H1) = &H1 Then
        Check1(8).Value = 1
    Else
        Check1(8).Value = 0
    End If
    
    If (UserRxdBuff(6) And &H2) = &H2 Then
        Check1(9).Value = 1
    Else
        Check1(9).Value = 0
    End If
    
    If (UserRxdBuff(6) And &H4) = &H4 Then
        Check1(10).Value = 1
    Else
        Check1(10).Value = 0
    End If
    
    If (UserRxdBuff(6) And &H8) = &H8 Then
        Check1(11).Value = 1
    Else
        Check1(11).Value = 0
    End If
    
    If (UserRxdBuff(6) And &H10) = &H10 Then
        Check1(12).Value = 1
    Else
        Check1(12).Value = 0
    End If
    If (UserRxdBuff(6) And &H20) = &H20 Then
        Check1(13).Value = 1
    Else
        Check1(13).Value = 0
    End If
    If (UserRxdBuff(6) And &H40) = &H40 Then
        Check1(14).Value = 1
    Else
        Check1(14).Value = 0
    End If
    If (UserRxdBuff(6) And &H80) = &H80 Then
        Check1(15).Value = 1
    Else
        Check1(15).Value = 0
    End If
End Sub
'Private Sub Check1_Click(Index As Integer)
Private Sub Check1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Bank14Data(1) As Byte
    Dim Msg, Style, Title, Help, Ctxt, Response
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Bank14Data(0) = 0
        Bank14Data(1) = 0
        If Check1(0) = 1 Then
            Bank14Data(0) = Bank14Data(0) Or &H1
        End If
        If Check1(1) = 1 Then
            Bank14Data(0) = Bank14Data(0) Or &H2
        End If
        If Check1(2) = 1 Then
            Bank14Data(0) = Bank14Data(0) Or &H4
        End If
        If Check1(3) = 1 Then
            Bank14Data(0) = Bank14Data(0) Or &H8
        End If
        If Check1(4) = 1 Then
            Bank14Data(0) = Bank14Data(0) Or &H10
        End If
        If Check1(5) = 1 Then
            Bank14Data(0) = Bank14Data(0) Or &H20
        End If
        If Check1(6) = 1 Then
            Bank14Data(0) = Bank14Data(0) Or &H40
        End If
        If Check1(7) = 1 Then
            Bank14Data(0) = Bank14Data(0) Or &H80
        End If
        
        If Check1(8) = 1 Then
            Bank14Data(1) = Bank14Data(1) Or &H1
        End If
        If Check1(9) = 1 Then
            Bank14Data(1) = Bank14Data(1) Or &H2
        End If
        If Check1(10) = 1 Then
            Bank14Data(1) = Bank14Data(1) Or &H4
        End If
        If Check1(11) = 1 Then
            Bank14Data(1) = Bank14Data(1) Or &H8
        End If
        If Check1(12) = 1 Then
            Bank14Data(1) = Bank14Data(1) Or &H10
        End If
        If Check1(13) = 1 Then
            Bank14Data(1) = Bank14Data(1) Or &H20
        End If
        If Check1(14) = 1 Then
            Bank14Data(1) = Bank14Data(1) Or &H40
        End If
        If Check1(15) = 1 Then
            Bank14Data(1) = Bank14Data(1) Or &H80
        End If
        Text12.Text = Right("0" + Hex(0), 2) '十六进制收数
        Text12.Text = Text12.Text & Right("0" + Hex(0), 2)  '十六进制收数
        Text12.Text = Text12.Text & Right("0" + Hex(Bank14Data(1)), 2)  '十六进制收数
        Text12.Text = Text12.Text & Right("0" + Hex(Bank14Data(0)), 2)  '十六进制收数
    End If
End Sub

Private Sub Check10_Click()

    If Check10.Value = 1 Then
        Text18.Enabled = False
    Else
        Text18.Enabled = True
    End If
End Sub

Private Sub Check2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer

    If Index < 5 Then
        For i = 0 To 4                  '速率
            Check2(i).Value = IIf(i = Index, 1, 0)
            If Check2(i).Value = 1 Then
                Combo3.ListIndex = i
            End If
        Next
    End If
    
    If Index = 5 Then
        Check2(5).Value = 1
        Check2(6).Value = 0
    End If
    If Index = 6 Then
        Check2(5).Value = 0
        Check2(6).Value = 1
    End If
    
    If Index = 7 Then
        Check2(7).Value = 1
        Check2(8).Value = 0
        Check2(9).Value = 0
    End If
    If Index = 8 Then
        Check2(7).Value = 0
        Check2(8).Value = 1
        Check2(9).Value = 0
    End If
    If Index = 9 Then
        Check2(7).Value = 0
        Check2(8).Value = 0
        Check2(9).Value = 1
    End If
    
    Call Bank4Data
End Sub
Private Sub Check3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim Msg, Style, Title, Help, Ctxt, Response
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        If WriteDel = 0 And Check3(11) = 1 Then
            Check3(12) = 0
            WriteDel = 1
            For i = 0 To 10
                Check3(i) = 1
            Next
        ElseIf WriteDel = 1 And Check3(12) = 1 Then
            Check3(11) = 0
            WriteDel = 0
            For i = 0 To 11
                Check3(i) = 0
            Next
        ElseIf Check3(11) = 0 And Check3(12) = 0 Then
            WriteDel = 0
        Else
            For i = 0 To 10
                Check3(i).Value = IIf(i = Index, 1, 0)
                If Check3(i).Value = 1 Then
                    Check3(11) = 0
                    Check3(12) = 0
                    WriteDel = 0
                End If
            Next
        End If
    End If
End Sub

Private Sub Command1_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text1.BackColor = &H80000005
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF1        '写
        UserBuff(2) = &H0         '块地址
        UserBuff(3) = &H1         '块地址
        cnt = 3
        s = Trim(Text1.Text)
        L = Len(s)
        Do While (L)
            Do Until IsHex(Left(s, 1)) <> 16
            L = L - 1
            If L = 0 Then Exit Do
            s = Right(s, L)
            Loop
            a = IsHex(Left(s, 1))
            L = L - 1
            s = Right(s, L)
            If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                a = a * 16 + IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
            End If
            cnt = cnt + 1
            UserBuff(cnt) = a
        Loop
        Call CRC16(UserBuff(), cnt)
        
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command10_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim cnt As Integer
    Dim L As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text2.Text = "00000000"
        Text2.BackColor = &H80000005
        UserBuff(0) = &HAA      '数据头
        UserBuff(1) = &HF0      '读
        Select Case Val(Combo3.ListIndex)   '读卡速率
            Case 0
                UserBuff(2) = &H8       '8rf
            Case 1
                UserBuff(2) = &H10      '16rf
            Case 2
                UserBuff(2) = &H20      '32rf
            Case 3
                UserBuff(2) = &H28      '40rf
            Case 4
                UserBuff(2) = &H40      '64rf
        End Select
        UserBuff(3) = &H0
        UserBuff(4) = &H8               '3块地址
        cnt = 4
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
     End If
End Sub

Private Sub Command11_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text3.Text = "00000000"
        Text3.BackColor = &H80000005
        UserBuff(0) = &HAA      '数据头
        UserBuff(1) = &HF0      '读
        
        Select Case Val(Combo3.ListIndex)   '读卡速率
            Case 0
                UserBuff(2) = &H8       '8rf
            Case 1
                UserBuff(2) = &H10      '16rf
            Case 2
                UserBuff(2) = &H20      '32rf
            Case 3
                UserBuff(2) = &H28      '40rf
            Case 4
                UserBuff(2) = &H40      '64rf
        End Select
        UserBuff(3) = &H0
        UserBuff(4) = &H20              '5块地址
        cnt = 4
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
     End If
End Sub

Private Sub Command12_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text4.Text = "00000000"
        Text4.BackColor = &H80000005
        UserBuff(0) = &HAA      '数据头
        UserBuff(1) = &HF0      '读
        
        Select Case Val(Combo3.ListIndex)   '读卡速率
            Case 0
                UserBuff(2) = &H8       '8rf
            Case 1
                UserBuff(2) = &H10      '16rf
            Case 2
                UserBuff(2) = &H20      '32rf
            Case 3
                UserBuff(2) = &H28      '40rf
            Case 4
                UserBuff(2) = &H40      '64rf
        End Select
        UserBuff(3) = &H0
        UserBuff(4) = &H40              '6块地址
        cnt = 4
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
     End If
End Sub

Private Sub Command13_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text5.Text = "00000000"
        Text5.BackColor = &H80000005
        UserBuff(0) = &HAA      '数据头
        UserBuff(1) = &HF0      '读
        
        Select Case Val(Combo3.ListIndex)   '读卡速率
            Case 0
                UserBuff(2) = &H8       '8rf
            Case 1
                UserBuff(2) = &H10      '16rf
            Case 2
                UserBuff(2) = &H20      '32rf
            Case 3
                UserBuff(2) = &H28      '40rf
            Case 4
                UserBuff(2) = &H40      '64rf
        End Select
        UserBuff(3) = &H0
        UserBuff(4) = &H80              '7块地址
        cnt = 4
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
     End If
End Sub

Private Sub Command14_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text6.Text = "00000000"
        Text6.BackColor = &H80000005
        UserBuff(0) = &HAA      '数据头
        UserBuff(1) = &HF0      '读
        
        Select Case Val(Combo3.ListIndex)   '读卡速率
            Case 0
                UserBuff(2) = &H8       '8rf
            Case 1
                UserBuff(2) = &H10      '16rf
            Case 2
                UserBuff(2) = &H20      '32rf
            Case 3
                UserBuff(2) = &H28      '40rf
            Case 4
                UserBuff(2) = &H40      '64rf
        End Select
        UserBuff(3) = &H1
        UserBuff(4) = &H0               '8块地址
        cnt = 4
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
     End If
End Sub

Private Sub Command15_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text7.Text = "00000000"
        Text7.BackColor = &H80000005
        UserBuff(0) = &HAA      '数据头
        UserBuff(1) = &HF0      '读
        
        Select Case Val(Combo3.ListIndex)   '读卡速率
            Case 0
                UserBuff(2) = &H8       '8rf
            Case 1
                UserBuff(2) = &H10      '16rf
            Case 2
                UserBuff(2) = &H20      '32rf
            Case 3
                UserBuff(2) = &H28      '40rf
            Case 4
                UserBuff(2) = &H40      '64rf
        End Select
        UserBuff(3) = &H2
        UserBuff(4) = &H0               '9块地址
        cnt = 4
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
     End If
End Sub

Private Sub Command16_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text8.Text = "00000000"
        Text8.BackColor = &H80000005
        UserBuff(0) = &HAA      '数据头
        UserBuff(1) = &HF0      '读
        
        Select Case Val(Combo3.ListIndex)   '读卡速率
            Case 0
                UserBuff(2) = &H8       '8rf
            Case 1
                UserBuff(2) = &H10      '16rf
            Case 2
                UserBuff(2) = &H20      '32rf
            Case 3
                UserBuff(2) = &H28      '40rf
            Case 4
                UserBuff(2) = &H40      '64rf
        End Select
        UserBuff(3) = &H4
        UserBuff(4) = &H0               '10块地址
        cnt = 4
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
     End If
End Sub

Private Sub Command17_Click()

    Dim L As Integer
    Dim TXD_Buff() As Byte
    
    On Error GoTo BLAK

    For L = 0 To 15
        Check1(L).Value = 0
    Next
    
    If Command17.Caption = "打 开" Then
        If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
            MSComm1.CommPort = Combo1.ListIndex + 1
            MSComm1.PortOpen = True
            Command17.Caption = "关 闭"
            Shape1.BackColor = RGB(0, 255, 0)
        Else
            MSComm1.CommPort = Combo1.ListIndex + 1
            MSComm1.PortOpen = True
            Command17.Caption = "关 闭"
            Shape1.BackColor = RGB(0, 255, 0)
            UserBuff(0) = &HAA       '数据头
            UserBuff(1) = &HFF       '握手
            UserBuff(2) = &HFF       '握手
            Call CRC16(UserBuff(), 2)
            UserBuff(3) = ReturnData(0)
            UserBuff(4) = ReturnData(1)
            UserBuff(5) = &HEE
            ReDim TXD_Buff(5)
            For L = 0 To 5
                    TXD_Buff(L) = UserBuff(L)
            Next L
            If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
                MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
            End If
            Combo1.Enabled = False
        End If
        Exit Sub
BLAK:
        MsgBox "串口不存在或被占用！", vbOKOnly, "提示信息"
    Else
        Command17.Caption = "打 开"
        Shape1.BackColor = RGB(255, 0, 0)
        MSComm1.PortOpen = False
        Combo1.Enabled = True
    End If
End Sub

Private Sub Command18_Click()

    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
'        For L = 0 To 18
'            Check2(L).Value = 0
'        Next
        Text15.Text = "00000000"
        Text15.BackColor = &HFFC0C0
        UserBuff(0) = &HAA      '数据头
        UserBuff(1) = &HF0      '读
        Select Case Val(Combo3.ListIndex)   '读卡速率
            Case 0
                UserBuff(2) = &H8       '8rf
            Case 1
                UserBuff(2) = &H10      '16rf
            Case 2
                UserBuff(2) = &H20      '32rf
            Case 3
                UserBuff(2) = &H28      '40rf
            Case 4
                UserBuff(2) = &H40      '64rf
        End Select
        UserBuff(3) = &H0
        UserBuff(4) = &H10               '4块地址
        cnt = 4
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
     End If
End Sub

Private Sub Command19_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text14.Text = "00000000"
        Text14.BackColor = &HFFFFC0
        UserBuff(0) = &HAA      '数据头
        UserBuff(1) = &HF0      '读
        
        Select Case Val(Combo3.ListIndex)   '读卡速率
            Case 0
                UserBuff(2) = &H8       '8rf
            Case 1
                UserBuff(2) = &H10      '16rf
            Case 2
                UserBuff(2) = &H20      '32rf
            Case 3
                UserBuff(2) = &H28      '40rf
            Case 4
                UserBuff(2) = &H40      '64rf
        End Select
        UserBuff(3) = &H0
        UserBuff(4) = &H2               '1块地址
        cnt = 4
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
     End If
End Sub

Private Sub Command2_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text2.BackColor = &H80000005
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF1        '写
        UserBuff(2) = &H0         '块地址
        UserBuff(3) = &H8         '块地址3
        cnt = 3
        s = Trim(Text2.Text)
        L = Len(s)
        Do While (L)
            Do Until IsHex(Left(s, 1)) <> 16
            L = L - 1
            If L = 0 Then Exit Do
            s = Right(s, L)
            Loop
            a = IsHex(Left(s, 1))
            L = L - 1
            s = Right(s, L)
            If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                a = a * 16 + IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
            End If
            cnt = cnt + 1
            UserBuff(cnt) = a
        Loop
        Call CRC16(UserBuff(), cnt)
        
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command20_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text13.Text = "00000000"
        Text13.BackColor = &H80000005
        UserBuff(0) = &HAA      '数据头
        UserBuff(1) = &HF0      '读
        
        Select Case Val(Combo3.ListIndex)   '读卡速率
            Case 0
                UserBuff(2) = &H8       '8rf
            Case 1
                UserBuff(2) = &H10      '16rf
            Case 2
                UserBuff(2) = &H20      '32rf
            Case 3
                UserBuff(2) = &H28      '40rf
            Case 4
                UserBuff(2) = &H40      '64rf
        End Select
        UserBuff(3) = &H80
        UserBuff(4) = &H0               '15块地址
        cnt = 4
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
     End If
End Sub

Private Sub Command21_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        For L = 0 To 15
            Check1(L).Value = 0
        Next
        Text12.Text = "00000000"
        Text12.BackColor = &H80000005
        UserBuff(0) = &HAA      '数据头
        UserBuff(1) = &HF0      '读
        
        Select Case Val(Combo3.ListIndex)   '读卡速率
            Case 0
                UserBuff(2) = &H8       '8rf
            Case 1
                UserBuff(2) = &H10      '16rf
            Case 2
                UserBuff(2) = &H20      '32rf
            Case 3
                UserBuff(2) = &H28      '40rf
            Case 4
                UserBuff(2) = &H40      '64rf
        End Select
        UserBuff(3) = &H40
        UserBuff(4) = &H0               '14块地址
        cnt = 4
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
     End If
End Sub

Private Sub Command22_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text11.Text = "00000000"
        Text11.BackColor = &H80000005
        UserBuff(0) = &HAA      '数据头
        UserBuff(1) = &HF0      '读
        
        Select Case Val(Combo3.ListIndex)   '读卡速率
            Case 0
                UserBuff(2) = &H8       '8rf
            Case 1
                UserBuff(2) = &H10      '16rf
            Case 2
                UserBuff(2) = &H20      '32rf
            Case 3
                UserBuff(2) = &H28      '40rf
            Case 4
                UserBuff(2) = &H40      '64rf
        End Select
        UserBuff(3) = &H20
        UserBuff(4) = &H0               '13块地址
        cnt = 4
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
     End If
End Sub

Private Sub Command23_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text10.Text = "00000000"
        Text10.BackColor = &H80000005
        UserBuff(0) = &HAA      '数据头
        UserBuff(1) = &HF0      '读
        
        Select Case Val(Combo3.ListIndex)   '读卡速率
            Case 0
                UserBuff(2) = &H8       '8rf
            Case 1
                UserBuff(2) = &H10      '16rf
            Case 2
                UserBuff(2) = &H20      '32rf
            Case 3
                UserBuff(2) = &H28      '40rf
            Case 4
                UserBuff(2) = &H40      '64rf
        End Select
        UserBuff(3) = &H10
        UserBuff(4) = &H0               '12块地址
        cnt = 4
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
     End If
End Sub

Private Sub Command24_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text9.Text = "00000000"
        Text9.BackColor = &H80000005
        UserBuff(0) = &HAA      '数据头
        UserBuff(1) = &HF0      '读
        
        Select Case Val(Combo3.ListIndex)   '读卡速率
            Case 0
                UserBuff(2) = &H8       '8rf
            Case 1
                UserBuff(2) = &H10      '16rf
            Case 2
                UserBuff(2) = &H20      '32rf
            Case 3
                UserBuff(2) = &H28      '40rf
            Case 4
                UserBuff(2) = &H40      '64rf
        End Select
        UserBuff(3) = &H8
        UserBuff(4) = &H0               '11块地址
        cnt = 4
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
     End If
End Sub

Private Sub Command25_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF2        '验证密码
        UserBuff(2) = &H0         '块地址
        UserBuff(3) = &H4         '块地址
        cnt = 3
        s = Trim(Text16.Text)
        L = Len(s)
        Do While (L)
            Do Until IsHex(Left(s, 1)) <> 16
            L = L - 1
            If L = 0 Then Exit Do
            s = Right(s, L)
            Loop
            a = IsHex(Left(s, 1))
            L = L - 1
            s = Right(s, L)
            If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                a = a * 16 + IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
            End If
            cnt = cnt + 1
            UserBuff(cnt) = a
        Loop
        Call CRC16(UserBuff(), cnt)
        
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command26_Click()

    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TX_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    ElseIf Check3(0) = 0 And Check3(1) = 0 And Check3(2) = 0 And Check3(3) = 0 And Check3(4) = 0 And Check3(5) = 0 And Check3(6) = 0 And Check3(7) = 0 And Check3(8) = 0 And Check3(9) = 0 And Check3(10) = 0 Then
        Msg = "请选择需要写入的块！！！"   '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF1        '写
        UserBuff(2) = &H0         '块地址
        UserBuff(3) = &H0         '块地址
        cnt = 3
        If Check3(0).Value = 1 Then
            UserBuff(3) = UserBuff(3) Or &H1 '0块
            s = Trim(Text1.Text)            '写入数据
            L = Len(s)
            Do While (L)
                Do Until IsHex(Left(s, 1)) <> 16
                L = L - 1
                If L = 0 Then Exit Do
                s = Right(s, L)
                Loop
                a = IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
                If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                    a = a * 16 + IsHex(Left(s, 1))
                    L = L - 1
                    s = Right(s, L)
                End If
                cnt = cnt + 1
                UserBuff(cnt) = a
            Loop
        End If
        If Check3(1).Value = 1 Then
            UserBuff(3) = UserBuff(3) Or &H8 '3块
            s = Trim(Text2.Text)            '写入数据
            L = Len(s)
            Do While (L)
                Do Until IsHex(Left(s, 1)) <> 16
                L = L - 1
                If L = 0 Then Exit Do
                s = Right(s, L)
                Loop
                a = IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
                If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                    a = a * 16 + IsHex(Left(s, 1))
                    L = L - 1
                    s = Right(s, L)
                End If
                cnt = cnt + 1
                UserBuff(cnt) = a
            Loop
        End If
        If Check3(2).Value = 1 Then
            UserBuff(3) = UserBuff(3) Or &H20 '5块
            s = Trim(Text3.Text)              '写入数据
            L = Len(s)
            Do While (L)
                Do Until IsHex(Left(s, 1)) <> 16
                L = L - 1
                If L = 0 Then Exit Do
                s = Right(s, L)
                Loop
                a = IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
                If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                    a = a * 16 + IsHex(Left(s, 1))
                    L = L - 1
                    s = Right(s, L)
                End If
                cnt = cnt + 1
                UserBuff(cnt) = a
            Loop
        End If
        
        If Check3(3).Value = 1 Then
            UserBuff(3) = UserBuff(3) Or &H40 '6块
            s = Trim(Text4.Text)              '写入数据
            L = Len(s)
            Do While (L)
                Do Until IsHex(Left(s, 1)) <> 16
                L = L - 1
                If L = 0 Then Exit Do
                s = Right(s, L)
                Loop
                a = IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
                If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                    a = a * 16 + IsHex(Left(s, 1))
                    L = L - 1
                    s = Right(s, L)
                End If
                cnt = cnt + 1
                UserBuff(cnt) = a
            Loop
        End If
        If Check3(4).Value = 1 Then
            UserBuff(3) = UserBuff(3) Or &H80 '7块
            s = Trim(Text5.Text)              '写入数据
            L = Len(s)
            Do While (L)
                Do Until IsHex(Left(s, 1)) <> 16
                L = L - 1
                If L = 0 Then Exit Do
                s = Right(s, L)
                Loop
                a = IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
                If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                    a = a * 16 + IsHex(Left(s, 1))
                    L = L - 1
                    s = Right(s, L)
                End If
                cnt = cnt + 1
                UserBuff(cnt) = a
            Loop
        End If
        If Check3(5).Value = 1 Then
            UserBuff(2) = UserBuff(2) Or &H1  '8块
            s = Trim(Text6.Text)              '写入数据
            L = Len(s)
            Do While (L)
                Do Until IsHex(Left(s, 1)) <> 16
                L = L - 1
                If L = 0 Then Exit Do
                s = Right(s, L)
                Loop
                a = IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
                If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                    a = a * 16 + IsHex(Left(s, 1))
                    L = L - 1
                    s = Right(s, L)
                End If
                cnt = cnt + 1
                UserBuff(cnt) = a
            Loop
        End If
        If Check3(6).Value = 1 Then
            UserBuff(2) = UserBuff(2) Or &H2  '9块
            s = Trim(Text7.Text)              '写入数据
            L = Len(s)
            Do While (L)
                Do Until IsHex(Left(s, 1)) <> 16
                L = L - 1
                If L = 0 Then Exit Do
                s = Right(s, L)
                Loop
                a = IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
                If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                    a = a * 16 + IsHex(Left(s, 1))
                    L = L - 1
                    s = Right(s, L)
                End If
                cnt = cnt + 1
                UserBuff(cnt) = a
            Loop
        End If
        If Check3(7).Value = 1 Then
            UserBuff(2) = UserBuff(2) Or &H4  '10块
            s = Trim(Text8.Text)              '写入数据
            L = Len(s)
            Do While (L)
                Do Until IsHex(Left(s, 1)) <> 16
                L = L - 1
                If L = 0 Then Exit Do
                s = Right(s, L)
                Loop
                a = IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
                If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                    a = a * 16 + IsHex(Left(s, 1))
                    L = L - 1
                    s = Right(s, L)
                End If
                cnt = cnt + 1
                UserBuff(cnt) = a
            Loop
        End If
        If Check3(8).Value = 1 Then
            UserBuff(2) = UserBuff(2) Or &H8  '11块
            s = Trim(Text9.Text)              '写入数据
            L = Len(s)
            Do While (L)
                Do Until IsHex(Left(s, 1)) <> 16
                L = L - 1
                If L = 0 Then Exit Do
                s = Right(s, L)
                Loop
                a = IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
                If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                    a = a * 16 + IsHex(Left(s, 1))
                    L = L - 1
                    s = Right(s, L)
                End If
                cnt = cnt + 1
                UserBuff(cnt) = a
            Loop
        End If
        If Check3(9).Value = 1 Then
            UserBuff(2) = UserBuff(2) Or &H10 '12块
            s = Trim(Text10.Text)             '写入数据
            L = Len(s)
            Do While (L)
                Do Until IsHex(Left(s, 1)) <> 16
                L = L - 1
                If L = 0 Then Exit Do
                s = Right(s, L)
                Loop
                a = IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
                If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                    a = a * 16 + IsHex(Left(s, 1))
                    L = L - 1
                    s = Right(s, L)
                End If
                cnt = cnt + 1
                UserBuff(cnt) = a
            Loop
        End If
        If Check3(10).Value = 1 Then
            UserBuff(2) = UserBuff(2) Or &H20  '13块
            s = Trim(Text11.Text)              '写入数据
            L = Len(s)
            Do While (L)
                Do Until IsHex(Left(s, 1)) <> 16
                L = L - 1
                If L = 0 Then Exit Do
                s = Right(s, L)
                Loop
                a = IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
                If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                    a = a * 16 + IsHex(Left(s, 1))
                    L = L - 1
                    s = Right(s, L)
                End If
                cnt = cnt + 1
                UserBuff(cnt) = a
            Loop
        End If
        Call CRC16(UserBuff(), cnt)
        
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TX_Buff(cnt + 3)
        
        For L = 0 To cnt + 3
            TX_Buff(L) = UserBuff(L)
        Next L
        
        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TX_Buff              '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command27_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Msg = "确定更新配置字吗？"          '定义消息文本
        Style = vbYesNo + vbDefaultButton2  ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then            ' 用户按下“是”
            Text15.BackColor = &H80000005
            UserBuff(0) = &HAA        '数据头
            UserBuff(1) = &HF1        '写
            UserBuff(2) = &H0         '块地址
            UserBuff(3) = &H10        '块地址 4
            cnt = 3
            
            s = Trim(Text15.Text)
            L = Len(s)
            Do While (L)
                Do Until IsHex(Left(s, 1)) <> 16
                L = L - 1
                If L = 0 Then Exit Do
                s = Right(s, L)
                Loop
                a = IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
                If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                    a = a * 16 + IsHex(Left(s, 1))
                    L = L - 1
                    s = Right(s, L)
                End If
                cnt = cnt + 1
                UserBuff(cnt) = a
            Loop
            
            Call CRC16(UserBuff(), cnt)
            
            UserBuff(cnt + 1) = ReturnData(0)
            UserBuff(cnt + 2) = ReturnData(1)
            UserBuff(cnt + 3) = &HEE
            ReDim TXD_Buff(cnt + 3)
            
            For L = 0 To cnt + 3
                TXD_Buff(L) = UserBuff(L)
            Next L
    
            If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
                MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
            End If
        End If
    End If
End Sub

Private Sub Command3_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text3.BackColor = &H80000005
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF1        '写
        UserBuff(2) = &H0         '块地址
        UserBuff(3) = &H20        '块地址5
        cnt = 3
        s = Trim(Text3.Text)
        L = Len(s)
        Do While (L)
            Do Until IsHex(Left(s, 1)) <> 16
            L = L - 1
            If L = 0 Then Exit Do
            s = Right(s, L)
            Loop
            a = IsHex(Left(s, 1))
            L = L - 1
            s = Right(s, L)
            If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                a = a * 16 + IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
            End If
            cnt = cnt + 1
            UserBuff(cnt) = a
        Loop
        Call CRC16(UserBuff(), cnt)
        
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command30_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text11.BackColor = &H80000005
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF1        '写
        UserBuff(2) = &H20        '块地址13
        UserBuff(3) = &H0         '块地址
        cnt = 3
        s = Trim(Text11.Text)
        L = Len(s)
        
        Do While (L)
            Do Until IsHex(Left(s, 1)) <> 16
            L = L - 1
            If L = 0 Then Exit Do
            s = Right(s, L)
            Loop
            a = IsHex(Left(s, 1))
            L = L - 1
            s = Right(s, L)
            If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                a = a * 16 + IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
            End If
            cnt = cnt + 1
            UserBuff(cnt) = a
        Loop
        Call CRC16(UserBuff(), cnt)
        
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command31_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text10.BackColor = &H80000005
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF1        '写
        UserBuff(2) = &H10        '块地址12
        UserBuff(3) = &H0         '块地址
        cnt = 3
        s = Trim(Text10.Text)
        L = Len(s)
        Do While (L)
            Do Until IsHex(Left(s, 1)) <> 16
            L = L - 1
            If L = 0 Then Exit Do
            s = Right(s, L)
            Loop
            a = IsHex(Left(s, 1))
            L = L - 1
            s = Right(s, L)
            If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                a = a * 16 + IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
            End If
            cnt = cnt + 1
            UserBuff(cnt) = a
        Loop
        Call CRC16(UserBuff(), cnt)
        
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command32_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text9.BackColor = &H80000005
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF1        '写
        UserBuff(2) = &H8         '块地址11
        UserBuff(3) = &H0         '块地址
        cnt = 3
        s = Trim(Text9.Text)
        L = Len(s)
        
        Do While (L)
            Do Until IsHex(Left(s, 1)) <> 16
            L = L - 1
            If L = 0 Then Exit Do
            s = Right(s, L)
            Loop
            a = IsHex(Left(s, 1))
            L = L - 1
            s = Right(s, L)
            If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                a = a * 16 + IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
            End If
            cnt = cnt + 1
            UserBuff(cnt) = a
        Loop
        
        Call CRC16(UserBuff(), cnt)
        
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command33_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        For L = 0 To 15
            Check1(L).Value = 0
        Next L
        Text1.Text = "00000000"
        Text2.Text = "00000000"
        Text3.Text = "00000000"
        Text4.Text = "00000000"
        Text5.Text = "00000000"
        Text6.Text = "00000000"
        Text7.Text = "00000000"
        Text8.Text = "00000000"
        Text9.Text = "00000000"
        Text10.Text = "00000000"
        Text11.Text = "00000000"
        Text12.Text = "00000000"
        Text13.Text = "00000000"
        Text14.Text = "00000000"
        Text15.Text = "00000000"
        Text1.BackColor = &H80000005
        Text2.BackColor = &H80000005
        Text3.BackColor = &H80000005
        Text4.BackColor = &H80000005
        Text5.BackColor = &H80000005
        Text6.BackColor = &H80000005
        Text7.BackColor = &H80000005
        Text8.BackColor = &H80000005
        Text9.BackColor = &H80000005
        Text10.BackColor = &H80000005
        Text11.BackColor = &H80000005
        Text12.BackColor = &H80000005
        Text13.BackColor = &H80000005
        Text14.BackColor = &H80000005
        Text15.BackColor = &H80000005
        
        UserBuff(0) = &HAA      '数据头
        UserBuff(1) = &HF0      '读
        Select Case Val(Combo3.ListIndex)   '读卡速率
            Case 0
                UserBuff(2) = &H8       '8rf
            Case 1
                UserBuff(2) = &H10      '16rf
            Case 2
                UserBuff(2) = &H20      '32rf
            Case 3
                UserBuff(2) = &H28      '40rf
            Case 4
                UserBuff(2) = &H40      '64rf
        End Select
        UserBuff(3) = &HFF              '不读密码块
        UserBuff(4) = &HFB
        cnt = 4
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
     End If
End Sub

Private Sub Command34_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Msg = "请牢记修改后的密码！！！"    '定义消息文本
        Style = vbYesNo + vbDefaultButton2  ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then            ' 用户按下“是”
            UserBuff(0) = &HAA        '数据头
            UserBuff(1) = &HF1        '修改密码
            UserBuff(2) = &H0         '块地址
            UserBuff(3) = &H4         '块地址
            cnt = 3
            s = Trim(Text19.Text)
            L = Len(s)
            Do While (L)
                Do Until IsHex(Left(s, 1)) <> 16
                L = L - 1
                If L = 0 Then Exit Do
                s = Right(s, L)
                Loop
                a = IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
                If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                    a = a * 16 + IsHex(Left(s, 1))
                    L = L - 1
                    s = Right(s, L)
                End If
                cnt = cnt + 1
                UserBuff(cnt) = a
            Loop
            Call CRC16(UserBuff(), cnt)
            
            UserBuff(cnt + 1) = ReturnData(0)
            UserBuff(cnt + 2) = ReturnData(1)
            UserBuff(cnt + 3) = &HEE
            ReDim TXD_Buff(cnt + 3)
            
            For L = 0 To cnt + 3
                TXD_Buff(L) = UserBuff(L)
            Next L
    
            If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
                MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
            End If
        End If
    End If
End Sub

Private Sub Command35_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF5        '休眠
        UserBuff(2) = &H0         '块地址
        UserBuff(3) = &H4         '块地址
        cnt = 3
        s = Trim(Text16.Text)
        L = Len(s)
        Do While (L)
            Do Until IsHex(Left(s, 1)) <> 16
            L = L - 1
            If L = 0 Then Exit Do
            s = Right(s, L)
            Loop
            a = IsHex(Left(s, 1))
            L = L - 1
            s = Right(s, L)
            If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                a = a * 16 + IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
            End If
            cnt = cnt + 1
            UserBuff(cnt) = a
        Loop
        Call CRC16(UserBuff(), cnt)
        
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command36_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Msg = "保护位不可逆，需谨慎修改！！！"        '定义消息文本
        Style = vbYesNo + vbDefaultButton2  ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then            ' 用户按下“是”
            UserBuff(0) = &HAA        '数据头
            UserBuff(1) = &HF1        '写
            UserBuff(2) = &H40        '块地址
            UserBuff(3) = &H0         '块地址 14
            cnt = 3
            s = Trim(Text12.Text)
            L = Len(s)
            Do While (L)
                Do Until IsHex(Left(s, 1)) <> 16
                L = L - 1
                If L = 0 Then Exit Do
                s = Right(s, L)
                Loop
                a = IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
                If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                    a = a * 16 + IsHex(Left(s, 1))
                    L = L - 1
                    s = Right(s, L)
                End If
                cnt = cnt + 1
                UserBuff(cnt) = a
            Loop
            Call CRC16(UserBuff(), cnt)
            
            UserBuff(cnt + 1) = ReturnData(0)
            UserBuff(cnt + 2) = ReturnData(1)
            UserBuff(cnt + 3) = &HEE
            ReDim TXD_Buff(cnt + 3)
            
            For L = 0 To cnt + 3
                TXD_Buff(L) = UserBuff(L)
            Next L
    
            If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
                MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
            End If
        End If
    End If
End Sub

Private Sub Command37_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Msg = "确定初始化卡片吗？"          '定义消息文本
        Style = vbYesNo + vbDefaultButton2  ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then            ' 用户按下“是”
            Check2(0).Value = 0
            Check2(1).Value = 0
            Check2(2).Value = 0
            Check2(3).Value = 0
            Check2(4).Value = 1
            Check2(5).Value = 1
            Check2(6).Value = 0
            Check2(7).Value = 1
            Check2(8).Value = 0
            Check2(9).Value = 0
            Check2(10).Value = 0
            Check2(11).Value = 1
            Check2(12).Value = 1
            Check2(13).Value = 0
            Check2(14).Value = 0
            Check2(15).Value = 0
            Check2(16).Value = 0
            Check2(17).Value = 0
            Check2(18).Value = 0
            Text15.Text = Right("000" + Hex(&H1805F), 8)
            UserBuff(0) = &HAA        '数据头
            UserBuff(1) = &HF1        '写
            UserBuff(2) = &H0         '块地址
            UserBuff(3) = &H10        '块地址 4
            cnt = 3
            
            s = Trim(Text15.Text)
            L = Len(s)
            Do While (L)
                Do Until IsHex(Left(s, 1)) <> 16
                L = L - 1
                If L = 0 Then Exit Do
                s = Right(s, L)
                Loop
                a = IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
                If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                    a = a * 16 + IsHex(Left(s, 1))
                    L = L - 1
                    s = Right(s, L)
                End If
                cnt = cnt + 1
                UserBuff(cnt) = a
            Loop
            
            Call CRC16(UserBuff(), cnt)
            
            UserBuff(cnt + 1) = ReturnData(0)
            UserBuff(cnt + 2) = ReturnData(1)
            UserBuff(cnt + 3) = &HEE
            ReDim TXD_Buff(cnt + 3)
            
            For L = 0 To cnt + 3
                TXD_Buff(L) = UserBuff(L)
            Next L
    
            If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
                MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
            End If
        End If
    End If
End Sub

Private Sub Command38_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Shape2.BackColor = &H80000005
        Text17.Text = "0000000000"
        Text17.BackColor = &H80000005
        If Check10.Value <> 1 Then
            Text18.Text = "0000000000"
            Text18.BackColor = &H80000005
        End If
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF3        '读4100卡
        cnt = 1
            
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
            
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L
    
        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command39_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Shape2.BackColor = &H80000005
        
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF4        '复制EM4100
        If Combo2.ListIndex = 0 Then
            UserBuff(2) = &H0               '基卡57卡
        Else
            UserBuff(2) = &H1               '基卡43卡
        End If
        cnt = 2
        s = Trim(Text18.Text)
        L = Len(s)
        Do While (L)
            Do Until IsHex(Left(s, 1)) <> 16
            L = L - 1
            If L = 0 Then Exit Do
            s = Right(s, L)
            Loop
            a = IsHex(Left(s, 1))
            L = L - 1
            s = Right(s, L)
            If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                a = a * 16 + IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
            End If
            cnt = cnt + 1
            UserBuff(cnt) = a
        Loop
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command4_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text4.BackColor = &H80000005
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF1        '写
        UserBuff(2) = &H0         '块地址
        UserBuff(3) = &H40        '块地址6
        cnt = 3
        s = Trim(Text4.Text)
        L = Len(s)
        Do While (L)
            Do Until IsHex(Left(s, 1)) <> 16
            L = L - 1
            If L = 0 Then Exit Do
            s = Right(s, L)
            Loop
            a = IsHex(Left(s, 1))
            L = L - 1
            s = Right(s, L)
            If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                a = a * 16 + IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
            End If
            cnt = cnt + 1
            UserBuff(cnt) = a
        Loop
        Call CRC16(UserBuff(), cnt)
        
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command40_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Shape2.BackColor = &H80000005
        Text17.Text = "0000000000"
        Text17.BackColor = &H80000005
        Text18.BackColor = &H80000005
        
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF3        '读4100卡
        cnt = 1
            
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
            
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L
    
        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command41_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Check2(0).Value = 0
        Check2(1).Value = 0
        Check2(2).Value = 0
        Check2(3).Value = 0
        Check2(4).Value = 1
        Check2(5).Value = 1
        Check2(6).Value = 0
        Check2(7).Value = 1
        Check2(8).Value = 0
        Check2(9).Value = 0
        Check2(10).Value = 0
        Check2(11).Value = 1
        Check2(12).Value = 1
        Check2(13).Value = 0
        Check2(14).Value = 0
        Check2(15).Value = 0
        Check2(16).Value = 0
        Check2(17).Value = 0
        Check2(18).Value = 0
        Text15.Text = Right("000" + Hex(&H1805F), 8)
    End If
End Sub

Private Sub Command5_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text5.BackColor = &H80000005
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF1        '写
        UserBuff(2) = &H0         '块地址
        UserBuff(3) = &H80        '块地址7
        cnt = 3
        s = Trim(Text5.Text)
        L = Len(s)
        Do While (L)
            Do Until IsHex(Left(s, 1)) <> 16
            L = L - 1
            If L = 0 Then Exit Do
            s = Right(s, L)
            Loop
            a = IsHex(Left(s, 1))
            L = L - 1
            s = Right(s, L)
            If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                a = a * 16 + IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
            End If
            cnt = cnt + 1
            UserBuff(cnt) = a
        Loop
        Call CRC16(UserBuff(), cnt)
        
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command6_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text6.BackColor = &H80000005
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF1        '写
        UserBuff(2) = &H1         '块地址8
        UserBuff(3) = &H0         '块地址
        cnt = 3
        
        s = Trim(Text6.Text)
        L = Len(s)
        Do While (L)
            Do Until IsHex(Left(s, 1)) <> 16
            L = L - 1
            If L = 0 Then Exit Do
            s = Right(s, L)
            Loop
            a = IsHex(Left(s, 1))
            L = L - 1
            s = Right(s, L)
            If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                a = a * 16 + IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
            End If
            cnt = cnt + 1
            UserBuff(cnt) = a
        Loop
        Call CRC16(UserBuff(), cnt)
        
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command7_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text7.BackColor = &H80000005
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF1        '写
        UserBuff(2) = &H2         '块地址9
        UserBuff(3) = &H0         '块地址
        cnt = 3
        
        s = Trim(Text7.Text)
        L = Len(s)
        
        Do While (L)
            Do Until IsHex(Left(s, 1)) <> 16
            L = L - 1
            If L = 0 Then Exit Do
            s = Right(s, L)
            Loop
            a = IsHex(Left(s, 1))
            L = L - 1
            s = Right(s, L)
            If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                a = a * 16 + IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
            End If
            cnt = cnt + 1
            UserBuff(cnt) = a
        Loop
        
        Call CRC16(UserBuff(), cnt)
        
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command8_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim s As String
    Dim L As Integer
    Dim cnt As Integer
    Dim TXD_Buff() As Byte

    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text8.BackColor = &H80000005
        UserBuff(0) = &HAA        '数据头
        UserBuff(1) = &HF1        '写
        UserBuff(2) = &H4         '块地址10
        UserBuff(3) = &H0         '块地址
        cnt = 3
        s = Trim(Text8.Text)
        L = Len(s)
        Do While (L)
            Do Until IsHex(Left(s, 1)) <> 16
            L = L - 1
            If L = 0 Then Exit Do
            s = Right(s, L)
            Loop
            a = IsHex(Left(s, 1))
            L = L - 1
            s = Right(s, L)
            If L <> 0 And IsHex(Left(s, 1)) <> 16 Then
                a = a * 16 + IsHex(Left(s, 1))
                L = L - 1
                s = Right(s, L)
            End If
            cnt = cnt + 1
            UserBuff(cnt) = a
        Loop
        
        Call CRC16(UserBuff(), cnt)
        
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff              '串口打开时将数据发送出去
        End If
    End If
End Sub

Private Sub Command9_Click()

    Dim Msg, Style, Title, Help, Ctxt, Response
    Dim cnt As Integer
    Dim L As Integer
    Dim TXD_Buff() As Byte
    
    If (MSComm1.PortOpen = False) = True Then   '判断串口是否打
        Msg = "请打开串口！！！"                '定义消息文本
        Style = vbDefaultButton2            ' 定义按钮+ vbCritical
        Title = "警告"                      ' 定义标题文本
        Help = "DEMO.HLP"                   ' 定义帮助文件
        Ctxt = 1000                         ' 定义帮助主题
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    Else
        Text1.Text = "00000000"
        Text1.BackColor = &H80000005
        UserBuff(0) = &HAA      '数据头
        UserBuff(1) = &HF0      '读
        
        Select Case Val(Combo3.ListIndex)   '读卡速率
            Case 0
                UserBuff(2) = &H8       '8rf
            Case 1
                UserBuff(2) = &H10      '16rf
            Case 2
                UserBuff(2) = &H20      '32rf
            Case 3
                UserBuff(2) = &H28      '40rf
            Case 4
                UserBuff(2) = &H40      '64rf
        End Select
        
        UserBuff(3) = &H0
        UserBuff(4) = &H1               '0块地址
        cnt = 4
        Call CRC16(UserBuff(), cnt)
        UserBuff(cnt + 1) = ReturnData(0)
        UserBuff(cnt + 2) = ReturnData(1)
        UserBuff(cnt + 3) = &HEE
        ReDim TXD_Buff(cnt + 3)
        For L = 0 To cnt + 3
            TXD_Buff(L) = UserBuff(L)
        Next L

        If (MSComm1.PortOpen = True) = True Then  '判断串口是否打开
            MSComm1.Output = TXD_Buff             '串口打开时将数据发送出去
        End If
     End If
End Sub

Private Sub Form_Load()

Dim i As Byte
Dim j As Byte

Shape1.BackColor = RGB(255, 0, 0)
Shape2.BackColor = RGB(255, 0, 0)
Combo2.ListIndex = 0
Check2(4).Value = 1
Check2(5).Value = 1
Check2(7).Value = 1
Check2(11).Value = 1
Check2(12).Value = 1
Combo3.ListIndex = 4
Combo2.ListIndex = 1
Text15.Text = Right("000" + Hex(&H1805F), 8)
Text19.Enabled = False
Command34.Enabled = False

With frmPort1

    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

    For i = 1 To 100
        On Error Resume Next '当运行发生错误时，控件转到紧接着发生错误的语句之后的语句，并在此继续运行
        MSComm1.CommPort = i
        MSComm1.PortOpen = True
        
        Select Case Err.Number
               Case 0                       '错误号为0(也就是没出错),
                    Combo1.ListIndex = i - 1
                    If j = 0 Then
                        j = i - 1
                    End If
                    MSComm1.PortOpen = False
                    Combo1.AddItem "COM" & Trim(i) & "  可用"
               Case 8005                    '错误号为8005,也就是端口被占用
                    Combo1.AddItem "COM" & Trim(i) & " 已经占用"
                    MSComm1.PortOpen = False
               Case Else
                    Combo1.AddItem "COM" & i
        End Select
        Err = 0               '将错误号置0. 注:Err.Number可以简写为Err ,2者等效
    Next
    Combo1.ListIndex = j
End With

End Sub


Private Sub MSComm1_OnComm()
    Dim Msg, Style, Title, Response
    Dim getBytes() As Byte           '用来从接收缓冲区读取数据
    Dim getLen As Integer
    Dim tmpi As Integer
    Dim i As Integer
    
    Select Case MSComm1.CommEvent
           Case comEvReceive
                MSComm1.RThreshold = 0   '关闭接收中断
                Sleep 100                '阻塞，等待接收完成
                MSComm1.RThreshold = 1
                getLen = MSComm1.InBufferCount '获取接收缓存的数据长度
                getBytes = MSComm1.Input       '取接收到的数据
                For tmpi = 0 To getLen - 1     '取接收到缓存数组
                    UserRxdBuff(tmpi) = getBytes(tmpi)
                Next tmpi
                
                If UserRxdBuff(0) = &HAA And UserRxdBuff(getLen - 1) = &HEE Then '判断数据头尾
                    Call CRC16(UserRxdBuff(), getLen - 4)            '计算CRC
                    If UserRxdBuff(getLen - 2) = ReturnData(0) And UserRxdBuff(getLen - 3) = ReturnData(1) Then   '判断CRC
                        Select Case UserRxdBuff(1)
                            Case &HF0    '读返回的数据
                                    If UserRxdBuff(2) = 1 Then  '成功
                                         Select Case UserRxdBuff(3)   '地址
                                            Case &H0        '块0
                                                    Text1.BackColor = RGB(0, 255, 0)
                                                    Text1.Text = ""
                                                    For i = 0 To 3
                                                        Text1.Text = Text1.Text & Right("0" + Hex(UserRxdBuff(i + 4)), 2)
                                                    Next i
                                            Case &H1        '块1   UID
                                                    Text14.BackColor = &HFFFFC0
                                                    Text14.Text = ""
                                                    For i = 0 To 3
                                                        Text14.Text = Text14.Text & Right("0" + Hex(UserRxdBuff(i + 4)), 2)
                                                    Next i
                                            Case &H2        '块2   密码
                                                    Text16.BackColor = &HC0E0FF
                                                    Text16.Text = ""
                                                    For i = 0 To 3
                                                        Text16.Text = Text16.Text & Right("0" + Hex(UserRxdBuff(i + 4)), 2)
                                                    Next i
                                            Case &H3        '块3
                                                    Text2.BackColor = RGB(0, 255, 0)
                                                    Text2.Text = ""
                                                    For i = 0 To 3
                                                        Text2.Text = Text2.Text & Right("0" + Hex(UserRxdBuff(i + 4)), 2)
                                                    Next i
                                            Case &H4        '块4   配置字
                                            
                                                    Text15.BackColor = &HFFC0C0
                                                    Text15.Text = ""
                                                    For i = 0 To 3
                                                        Text15.Text = Text15.Text & Right("0" + Hex(UserRxdBuff(i + 4)), 2)
                                                    Next i
                                                    Call UartBank4Data   '配置块勾选
                                            Case &H5        '块5
                                                    Text3.BackColor = RGB(0, 255, 0)
                                                    Text3.Text = ""
                                                    For i = 0 To 3
                                                        Text3.Text = Text3.Text & Right("0" + Hex(UserRxdBuff(i + 4)), 2)
                                                    Next i
                                            Case &H6        '块6
                                                    Text4.BackColor = RGB(0, 255, 0)
                                                    Text4.Text = ""
                                                    For i = 0 To 3
                                                        Text4.Text = Text4.Text & Right("0" + Hex(UserRxdBuff(i + 4)), 2)
                                                    Next i
                                            Case &H7        '块7
                                                    Text5.BackColor = RGB(0, 255, 0)
                                                    Text5.Text = ""
                                                    For i = 0 To 3
                                                        Text5.Text = Text5.Text & Right("0" + Hex(UserRxdBuff(i + 4)), 2)
                                                    Next i
                                            Case &H8        '块8
                                                    Text6.BackColor = RGB(0, 255, 0)
                                                    Text6.Text = ""
                                                    For i = 0 To 3
                                                        Text6.Text = Text6.Text & Right("0" + Hex(UserRxdBuff(i + 4)), 2)
                                                    Next i
                                            Case &H9        '块9
                                                    Text7.BackColor = RGB(0, 255, 0)
                                                    Text7.Text = ""
                                                    For i = 0 To 3
                                                        Text7.Text = Text7.Text & Right("0" + Hex(UserRxdBuff(i + 4)), 2)
                                                    Next i
                                            Case &HA        '块10
                                                    Text8.BackColor = RGB(0, 255, 0)
                                                    Text8.Text = ""
                                                    For i = 0 To 3
                                                        Text8.Text = Text8.Text & Right("0" + Hex(UserRxdBuff(i + 4)), 2)
                                                    Next i
                                            Case &HB        '块11
                                                    Text9.BackColor = RGB(0, 255, 0)
                                                    Text9.Text = ""
                                                    For i = 0 To 3
                                                        Text9.Text = Text9.Text & Right("0" + Hex(UserRxdBuff(i + 4)), 2)
                                                    Next i
                                            Case &HC        '块12
                                                    Text10.BackColor = RGB(0, 255, 0)
                                                    Text10.Text = ""
                                                    For i = 0 To 3
                                                        Text10.Text = Text10.Text & Right("0" + Hex(UserRxdBuff(i + 4)), 2)
                                                    Next i
                                            Case &HD        '块13
                                                    Text11.BackColor = RGB(0, 255, 0)
                                                    Text11.Text = ""
                                                    For i = 0 To 3
                                                        Text11.Text = Text11.Text & Right("0" + Hex(UserRxdBuff(i + 4)), 2)
                                                    Next i
                                            Case &HE        '块14  保护字1
                                                    Text12.BackColor = RGB(0, 255, 0)
                                                    Text12.Text = ""
                                                    For i = 0 To 3
                                                        Text12.Text = Text12.Text & Right("0" + Hex(UserRxdBuff(i + 4)), 2)
                                                    Next i
                                                    If UserRxdBuff(6) <> 0 Or UserRxdBuff(7) <> 0 Then
                                                          Call UartBank14Data     '锁勾选
                                                    End If
                                            Case &HF        '块15  保护字2
                                                    Text13.BackColor = RGB(0, 255, 0)
                                                    Text13.Text = ""
                                                    For i = 0 To 3
                                                        Text13.Text = Text13.Text & Right("0" + Hex(UserRxdBuff(i + 4)), 2)
                                                    Next i
                                                    If UserRxdBuff(6) <> 0 Or UserRxdBuff(7) <> 0 Then
                                                         Call UartBank14Data     '锁勾选
                                                    End If
                                         End Select
                                    Else
                                         Select Case UserRxdBuff(3)   '地址
                                                Case &H0
                                                        Text1.Text = "00000000"   '扇区0
                                                        Text1.BackColor = &HC0C0FF
                                                Case &H1
                                                        Text14.Text = "00000000"  '扇区1 UID
                                                        Text14.BackColor = &HC0C0FF
                                                Case &H2
                                                Case &H3
                                                        Text2.Text = "00000000"   '扇区3
                                                        Text2.BackColor = &HC0C0FF
                                                Case &H4
                                                        Text15.Text = "00000000"  '扇区4 配置块
                                                        Text15.BackColor = &HC0C0FF
                                                Case &H5
                                                        Text3.Text = "00000000"   '扇区5
                                                        Text3.BackColor = &HC0C0FF
                                                Case &H6
                                                        Text4.Text = "00000000"   '扇区6
                                                        Text4.BackColor = &HC0C0FF
                                                Case &H7
                                                        Text5.Text = "00000000"   '扇区7
                                                        Text5.BackColor = &HC0C0FF
                                                Case &H8
                                                        Text6.Text = "00000000"   '扇区8
                                                        Text6.BackColor = &HC0C0FF
                                                Case &H9
                                                        Text7.Text = "00000000"   '扇区9
                                                        Text7.BackColor = &HC0C0FF
                                                Case &HA
                                                        Text8.Text = "00000000"   '扇区10
                                                        Text8.BackColor = &HC0C0FF
                                                Case &HB
                                                        Text9.Text = "00000000"   '扇区11
                                                        Text9.BackColor = &HC0C0FF
                                                Case &HC
                                                        Text10.Text = "00000000"  '扇区12
                                                        Text10.BackColor = &HC0C0FF
                                                Case &HD
                                                        Text11.Text = "00000000"  '扇区13
                                                        Text11.BackColor = &HC0C0FF
                                                Case &HE
                                                        Text12.Text = "00000000"   '扇区14
                                                        Text12.BackColor = &HC0C0FF
                                                Case &HF
                                                        Text13.Text = "00000000"   '扇区15
                                                        Text13.BackColor = &HC0C0FF
                                        End Select
                                    End If
                            Case &HF1    '写返回的数据
                                    If UserRxdBuff(2) = 1 Then  '成功
                                        Select Case UserRxdBuff(3)   '地址
                                                Case &H0
                                                        Text1.BackColor = RGB(0, 255, 0)  '扇区0
                                                Case &H1
                                                        Text14.BackColor = &HFFFFC0       '扇区1 UID
                                                Case &H2
                                                        Text19.BackColor = &H80FF80      '扇区2 密码块 说明修改密码成功
                                                        Text16.Text = Text19.Text        '复制到密码区
                                                        Text19.Enabled = False
                                                        Command34.Enabled = False
                                                        Msg = "修改成功"                 '定义消息文本
                                                        Style = vbYes
                                                        Title = "提示"                    ' 定义标题文本
                                                        Response = MsgBox(Msg, Style, Title)
                                                Case &H3
                                                        Text2.BackColor = RGB(0, 255, 0)  '扇区3
                                                Case &H4
                                                        Text15.BackColor = RGB(0, 255, 0) '扇区4 配置块
                                                Case &H5
                                                        Text3.BackColor = RGB(0, 255, 0)  '扇区5
                                                Case &H6
                                                        Text4.BackColor = RGB(0, 255, 0)  '扇区6
                                                Case &H7
                                                        Text5.BackColor = RGB(0, 255, 0)  '扇区7
                                                Case &H8
                                                        Text6.BackColor = RGB(0, 255, 0)  '扇区8
                                                Case &H9
                                                        Text7.BackColor = RGB(0, 255, 0)  '扇区9
                                                Case &HA
                                                        Text8.BackColor = RGB(0, 255, 0)  '扇区10
                                                Case &HB
                                                        Text9.BackColor = RGB(0, 255, 0)  '扇区11
                                                Case &HC
                                                        Text10.BackColor = RGB(0, 255, 0) '扇区12
                                                Case &HD
                                                        Text11.BackColor = RGB(0, 255, 0) '扇区13
                                                Case &HE
                                                        Text12.BackColor = RGB(0, 255, 0) '扇区14
                                                Case &HF
                                                        Text13.BackColor = RGB(0, 255, 0) '扇区15
                                        End Select
                                    Else
                                        Select Case UserRxdBuff(3)   '地址
                                                Case 0
                                                        Text1.BackColor = &HC0C0FF  '扇区0
                                                Case 1
                                                        Text14.BackColor = &HC0C0FF   '扇区1 UID
                                                Case 2
                                                        Text16.BackColor = &HC0C0FF '扇区2 密码块 修改失败
                                                        Text19.Enabled = False
                                                        Command34.Enabled = False
                                                        Msg = "修改失败"                  '定义消息文本
                                                        Style = vbYes
                                                        Title = "提示"                    ' 定义标题文本
                                                        Response = MsgBox(Msg, Style, Title)
                                                Case 3
                                                        Text2.BackColor = &HC0C0FF  '扇区3
                                                Case 4
                                                        Text15.BackColor = &HC0C0FF '扇区4 配置块
                                                Case 5
                                                        Text3.BackColor = &HC0C0FF  '扇区5
                                                Case 6
                                                        Text4.BackColor = &HC0C0FF  '扇区6
                                                Case 7
                                                        Text5.BackColor = &HC0C0FF  '扇区7
                                                Case 8
                                                        Text6.BackColor = &HC0C0FF  '扇区8
                                                Case 9
                                                        Text7.BackColor = &HC0C0FF  '扇区9
                                                Case 10
                                                        Text8.BackColor = &HC0C0FF  '扇区10
                                                Case 11
                                                        Text9.BackColor = &HC0C0FF  '扇区11
                                                Case 12
                                                        Text10.BackColor = &HC0C0FF '扇区12
                                                Case 13
                                                        Text11.BackColor = &HC0C0FF '扇区13
                                                Case 14
                                                        Text12.BackColor = &HC0C0FF '扇区14
                                                        Text13.BackColor = &HC0C0FF '扇区15
                                                Case 15
                                                        Text12.BackColor = &HC0C0FF '扇区14
                                                        Text13.BackColor = &HC0C0FF '扇区15
                                        End Select
                                    End If
                            Case &HF2       '登录 验证密码
                                    If UserRxdBuff(2) = 1 Then  '成功
                                        Text16.BackColor = &H80FF80      '扇区2 密码块3
                                        Text19.Enabled = True
                                        Command34.Enabled = True
                                        Msg = "登录成功"                 '定义消息文本
                                        Style = vbYes
                                        Title = "提示"               ' 定义标题文本
                                        Response = MsgBox(Msg, Style, Title)
                                        PassTimeS = 50
                                    Else  '失败   '
                                        Text16.BackColor = &HFF&     '密码错误 扇区2 密码块
                                        Text19.Enabled = False
                                        Command34.Enabled = False
                                        Msg = "密码错误"             '定义消息文本
                                        Style = vbYes
                                        Title = "警告"               ' 定义标题文本
                                        Response = MsgBox(Msg, Style, Title)
                                    End If
                            Case &HF3        '读EM4100返回的数据
                                    If UserRxdBuff(2) = 1 Then  '成功
                                        Shape2.BackColor = RGB(0, 255, 0)
                                        Text17.BackColor = RGB(0, 255, 0)
                                        Text17.Text = ""
                                        For i = 0 To 4
                                            Text17.Text = Text17.Text & Right("0" + Hex(UserRxdBuff(i + 3)), 2)   '扇区2
                                        Next i
                                    Else
                                        Shape2.BackColor = RGB(255, 0, 0)
                                        Text17.Text = "0000000000"
                                        Text17.BackColor = &HC0C0FF
                                    End If
                                    
                                    If Check10.Value = 0 Then
                                        Text18.Text = Text17.Text
                                    End If
                                    
                            Case &HF4       '复制EM4100返回的数据
                                    If UserRxdBuff(2) = 1 Then  '成功
                                        Shape2.BackColor = RGB(0, 255, 0)
                                        Text18.BackColor = RGB(0, 255, 0)
                                        If Check10.Value = 1 Then   '自增1
                                            Call EM4100_Updata
                                        End If
                                    Else
                                        Shape2.BackColor = RGB(255, 0, 0)
                                        Text18.BackColor = &HC0C0FF
                                    End If
                            Case &HF5
                        End Select
                    End If
                End If
                
    End Select
End Sub


Private Sub Text1_Change()
    If Len(Text1.Text) = 9 Then
        Text1.Text = ""
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Text10_Change()
    If Len(Text10.Text) = 9 Then
        Text10.Text = ""
    End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Text11_Change()
    If Len(Text11.Text) = 9 Then
        Text11.Text = ""
    End If
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Text12_Change()
    If Len(Text12.Text) = 9 Then
        Text12.Text = ""
    End If
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Text13_Change()
If Len(Text13.Text) = 9 Then
Text13.Text = ""
End If
End Sub

Private Sub Text14_Change()
    If Len(Text14.Text) = 9 Then
        Text14.Text = ""
    End If
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Text15_Change()
    If Len(Text15.Text) = 9 Then
        Text15.Text = ""
    End If
End Sub
Private Sub Text15_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Text16_Change()
    If Len(Text16.Text) = 9 Then
        Text16.Text = ""
        Text16.BackColor = &HC0E0FF
    End If
End Sub
Private Sub Text16_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub


Private Sub Text17_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Text18_Change()
    If Len(Text18.Text) = 11 Then
         Text18.Text = ""
    End If
End Sub
Private Sub Text18_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Text19_Change()
    If Len(Text19.Text) = 9 Then
        Text2.Text = ""
    End If
End Sub
Private Sub Text19_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Text2_Change()
    If Len(Text2.Text) = 9 Then
        Text2.Text = ""
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Text3_Change()
    If Len(Text3.Text) = 9 Then
        Text3.Text = ""
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Text4_Change()
    If Len(Text4.Text) = 9 Then
        Text4.Text = ""
    End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Text5_Change()
    If Len(Text5.Text) = 9 Then
        Text5.Text = ""
    End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Text6_Change()
    If Len(Text6.Text) = 9 Then
        Text6.Text = ""
    End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Text7_Change()
    If Len(Text7.Text) = 9 Then
        Text7.Text = ""
    End If
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Text8_Change()
    If Len(Text8.Text) = 9 Then
        Text8.Text = ""
    End If
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Text9_Change()
    If Len(Text9.Text) = 9 Then
        Text9.Text = ""
    End If
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
  Dim temp As String
  If (KeyAscii <= vbKey9 And KeyAscii >= vbKey0) Or (KeyAscii <= 102 And KeyAscii >= 97) Or (KeyAscii <= vbKeyF And KeyAscii >= vbKeyA) Or KeyAscii = 8 Then
     If KeyAscii <= 102 And KeyAscii >= 97 Then
        KeyAscii = KeyAscii - 32
     End If
     Else
      KeyAscii = 0
  End If
End Sub

Private Sub Timer1_Timer()
    If PassTimeS <> 0 Then
        PassTimeS = PassTimeS - 1
        If PassTimeS = 0 Then
            Command34.Enabled = False
            Text19.Enabled = False
        End If
    End If
End Sub
