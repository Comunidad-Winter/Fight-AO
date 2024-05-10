VERSION 5.00
Begin VB.Form frmOpciones 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3045
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdMsn 
      BackColor       =   &H000080FF&
      Caption         =   "Mensaje de Msn Activado"
      Height          =   345
      Left            =   120
      MouseIcon       =   "frmOpciones.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Width           =   2790
   End
   Begin VB.CommandButton cmdSound 
      BackColor       =   &H000080FF&
      Caption         =   "Sonidos"
      Height          =   345
      Left            =   105
      MouseIcon       =   "frmOpciones.frx":02A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   2790
   End
   Begin VB.CommandButton cmdMusica 
      BackColor       =   &H000080FF&
      Caption         =   "Musica"
      Height          =   345
      Left            =   105
      MouseIcon       =   "frmOpciones.frx":03F6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   2790
   End
   Begin VB.CommandButton cmdForo 
      BackColor       =   &H000080FF&
      Caption         =   "FORO"
      Height          =   345
      Index           =   6
      Left            =   105
      MouseIcon       =   "frmOpciones.frx":0548
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   2790
   End
   Begin VB.CommandButton CmdMapa 
      BackColor       =   &H000080FF&
      Caption         =   "MAPA"
      Height          =   345
      Index           =   5
      Left            =   105
      MouseIcon       =   "frmOpciones.frx":069A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   2790
   End
   Begin VB.CommandButton cmdAlphaB 
      BackColor       =   &H000080FF&
      Caption         =   "Alpha Blending Activado"
      Height          =   345
      Left            =   105
      MouseIcon       =   "frmOpciones.frx":07EC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   2790
   End
   Begin VB.CommandButton CmdUclick 
      BackColor       =   &H000080FF&
      Caption         =   "U+Click Boton derecho Activado"
      Height          =   345
      Left            =   105
      MouseIcon       =   "frmOpciones.frx":093E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   2790
   End
   Begin VB.CommandButton Command3 
      Caption         =   "a"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H000080FF&
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   105
      MouseIcon       =   "frmOpciones.frx":0A90
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   2790
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   2775
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAlphaB_Click()
If ConAlfaB = True Then
ConAlfaB = False
cmdAlphaB.Caption = "AlphaBlending Desactivado"
Else
ConAlfaB = True
cmdAlphaB.Caption = "AlphaBlending Activado"
End If
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub CmdMapa_Click(index As Integer)
Call frmMapa.Show(vbModeless, frmMain)
End Sub

Private Sub cmdMsn_Click()
Call SetMusicInfo("", "", "", "Jugando Twist-AO", , False)
End Sub

Private Sub cmdMusica_Click()
        If Musica Then
            Musica = False
            cmdMusica.Caption = "Musica Desactivada"
            Audio.StopMidi
        Else
            Musica = True
            cmdMusica.Caption = "Musica Activada"
            Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
        End If
End Sub
Private Sub cmdSound_Click()
If Sound Then
            Sound = False
            cmdSound.Caption = "Sonidos Desactivados"
            Call Audio.StopWave
            RainBufferIndex = 0
            frmMain.IsPlaying = PlayLoop.plNone
        Else
            Sound = True
            cmdSound.Caption = "Sonidos Activados"
        End If
End Sub

Private Sub CmdUclick_Click()
If Uclickear = True Then
Uclickear = False
CmdUclick.Caption = "U+Click Boton derecho Desactivado"
Else
Uclickear = True
CmdUclick.Caption = "U+Click Boton derecho Activado"
End If
End Sub

Private Sub Form_Load()
    If Musica Then
        cmdMusica.Caption = "Musica Activada"
    Else
        cmdMusica.Caption = "Musica Desactivada"
    End If
    
    If Sound Then
        cmdSound.Caption = "Sonidos Activados"
    Else
        cmdSound.Caption = "Sonidos Desactivados"
    End If
End Sub
