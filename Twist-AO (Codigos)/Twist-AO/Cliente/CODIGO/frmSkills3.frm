VERSION 5.00
Begin VB.Form frmSkills3 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignar Skills Points"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4185
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   4185
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "ACEPTAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7680
      Width           =   3690
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   630
      TabIndex        =   42
      Top             =   420
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   630
      TabIndex        =   41
      Top             =   765
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   630
      TabIndex        =   40
      Top             =   1110
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   630
      TabIndex        =   39
      Top             =   1455
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   630
      TabIndex        =   38
      Top             =   1800
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   5
      Left            =   630
      TabIndex        =   37
      Top             =   2145
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   6
      Left            =   630
      TabIndex        =   36
      Top             =   2445
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   7
      Left            =   630
      TabIndex        =   35
      Top             =   2835
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   8
      Left            =   630
      TabIndex        =   34
      Top             =   3195
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   9
      Left            =   630
      TabIndex        =   33
      Top             =   3540
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   10
      Left            =   630
      TabIndex        =   32
      Top             =   3885
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   11
      Left            =   630
      TabIndex        =   31
      Top             =   4230
      Width           =   2205
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   30
      Top             =   390
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   29
      Top             =   735
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   28
      Top             =   1095
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   27
      Top             =   1440
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   26
      Top             =   1800
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   25
      Top             =   2160
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   3240
      TabIndex        =   24
      Top             =   2490
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   3240
      TabIndex        =   23
      Top             =   2835
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   3240
      TabIndex        =   22
      Top             =   3180
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   3240
      TabIndex        =   21
      Top             =   3525
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   11
      Left            =   3240
      TabIndex        =   20
      Top             =   3885
      Width           =   285
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   19
      Top             =   4230
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   0
      Left            =   3600
      Top             =   390
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   2
      Left            =   3600
      Top             =   750
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   3
      Left            =   2820
      Top             =   780
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   4
      Left            =   3600
      Top             =   1095
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   5
      Left            =   2820
      Top             =   1125
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   6
      Left            =   3600
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   7
      Left            =   2820
      Top             =   1470
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   8
      Left            =   3600
      Top             =   1785
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   9
      Left            =   2820
      Top             =   1815
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   10
      Left            =   3600
      Top             =   2130
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   11
      Left            =   2820
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   12
      Left            =   3600
      Top             =   2475
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   13
      Left            =   2820
      Top             =   2505
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   14
      Left            =   3600
      Top             =   2820
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   15
      Left            =   2820
      Top             =   2850
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   16
      Left            =   3600
      Top             =   3180
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   17
      Left            =   2820
      Top             =   3180
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   18
      Left            =   3600
      Top             =   3525
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   19
      Left            =   2820
      Top             =   3525
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   20
      Left            =   3600
      Top             =   3870
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   21
      Left            =   2820
      Top             =   3870
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   22
      Left            =   3600
      Top             =   4215
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   23
      Left            =   2820
      Top             =   4215
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   24
      Left            =   3600
      Top             =   4560
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   25
      Left            =   2820
      Top             =   4560
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   13
      Left            =   3240
      TabIndex        =   18
      Top             =   4560
      Width           =   285
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   12
      Left            =   630
      TabIndex        =   17
      Top             =   4575
      Width           =   2205
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   26
      Left            =   3600
      Top             =   4905
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   27
      Left            =   2820
      Top             =   4905
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   14
      Left            =   3240
      TabIndex        =   16
      Top             =   4920
      Width           =   285
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   13
      Left            =   630
      TabIndex        =   15
      Top             =   4920
      Width           =   2205
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   28
      Left            =   3600
      Top             =   5250
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   29
      Left            =   2820
      Top             =   5250
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   15
      Left            =   3240
      TabIndex        =   14
      Top             =   5265
      Width           =   285
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   14
      Left            =   630
      TabIndex        =   13
      Top             =   5265
      Width           =   2205
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   30
      Left            =   3600
      Top             =   5595
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   31
      Left            =   2820
      Top             =   5595
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   16
      Left            =   3240
      TabIndex        =   12
      Top             =   5640
      Width           =   285
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   15
      Left            =   630
      TabIndex        =   11
      Top             =   5610
      Width           =   2205
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   32
      Left            =   3600
      Top             =   5940
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   33
      Left            =   2820
      Top             =   5940
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   17
      Left            =   3240
      TabIndex        =   10
      Top             =   5970
      Width           =   285
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   16
      Left            =   630
      TabIndex        =   9
      Top             =   5955
      Width           =   2205
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   34
      Left            =   3600
      Top             =   6285
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   35
      Left            =   2820
      Top             =   6285
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   18
      Left            =   3240
      TabIndex        =   8
      Top             =   6330
      Width           =   285
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   17
      Left            =   630
      TabIndex        =   7
      Top             =   6300
      Width           =   2205
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   1
      Left            =   2820
      Top             =   435
      Width           =   345
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   18
      Left            =   630
      TabIndex        =   6
      Top             =   6645
      Width           =   2205
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   19
      Left            =   3240
      TabIndex        =   5
      Top             =   6675
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   36
      Left            =   3600
      Top             =   6630
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   37
      Left            =   2820
      Top             =   6645
      Width           =   345
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   19
      Left            =   630
      TabIndex        =   4
      Top             =   6990
      Width           =   2205
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   20
      Left            =   3240
      TabIndex        =   3
      Top             =   7035
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   38
      Left            =   3600
      Top             =   6975
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   39
      Left            =   2820
      Top             =   6945
      Width           =   345
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   20
      Left            =   630
      TabIndex        =   2
      Top             =   7350
      Width           =   2205
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   21
      Left            =   3240
      TabIndex        =   1
      Top             =   7335
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   40
      Left            =   3600
      Top             =   7320
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   41
      Left            =   2820
      Top             =   7290
      Width           =   345
   End
   Begin VB.Label puntos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1680
      TabIndex        =   0
      Top             =   60
      Width           =   660
   End
End
Attribute VB_Name = "frmSkills3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Private Sub Command1_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

Dim indice
If index Mod 2 = 0 Then
    If Alocados > 0 Then
        indice = index \ 2 + 1
        If indice > NUMSKILLS Then indice = NUMSKILLS
        If Val(text1(indice).Caption) < MAXSKILLPOINTS Then
            text1(indice).Caption = Val(text1(indice).Caption) + 1
            flags(indice) = flags(indice) + 1
            Alocados = Alocados - 1
        End If
            
    End If
Else
    If Alocados < SkillPoints Then
        
        indice = index \ 2 + 1
        If Val(text1(indice).Caption) > 0 And flags(indice) > 0 Then
            text1(indice).Caption = Val(text1(indice).Caption) - 1
            flags(indice) = flags(indice) - 1
            Alocados = Alocados + 1
        End If
    End If
End If

puntos.Caption = "Puntos:" & Alocados
End Sub

Private Sub Command2_Click()
Dim i As Integer
    Dim cad As String
    
    For i = 1 To NUMSKILLS
        cad = cad & flags(i) & ","
        'Actualizamos nuestros datos locales
        UserSkills(i) = Val(text1(i).Caption)
    Next i
    
    SendData "SKSE" & cad
    If Alocados = 0 Then frmMain.Label1.Visible = False
    SkillPoints = Alocados
    Unload Me
End Sub

Private Sub Form_Deactivate()
'Me.Visible = False
End Sub

Private Sub Form_Load()

'Nombres de los skills

Dim l
Dim i As Integer
i = 1
For Each l In Label2
    l.Caption = SkillsNames(i)
    l.AutoSize = True
    i = i + 1
Next
i = 0

'Flags para saber que skills se modificaron
ReDim flags(1 To NUMSKILLS)


'Cargamos el jpg correspondiente
For i = 0 To NUMSKILLS * 2 - 1
    If i Mod 2 = 0 Then
        command1(i).Picture = LoadPicture(App.Path & "\Graficos\botonmas.bmp")
    Else
        command1(i).Picture = LoadPicture(App.Path & "\Graficos\botonmenos.bmp")
    End If
Next

'Alocados = SkillPoints
End Sub