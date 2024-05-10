VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   480
      Top             =   5400
   End
   Begin VB.TextBox PasswordTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   1875
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   7260
      Width           =   3330
   End
   Begin VB.TextBox NombreTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   1875
      TabIndex        =   4
      Top             =   6390
      Width           =   3330
   End
   Begin VB.ListBox lst_servers 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   450
      ItemData        =   "frmConnect.frx":000C
      Left            =   1080
      List            =   "frmConnect.frx":0013
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   2160
      TabIndex        =   0
      Text            =   "7666"
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   2640
      TabIndex        =   2
      Text            =   "localhost"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1800
      TabIndex        =   6
      Top             =   7680
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   5880
      Top             =   6000
      Width           =   2775
   End
   Begin VB.Image imgServEspana 
      Height          =   315
      Left            =   3360
      MousePointer    =   99  'Custom
      Top             =   360
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgServArgentina 
      Height          =   435
      Left            =   4080
      MousePointer    =   99  'Custom
      Top             =   240
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgGetPass 
      Height          =   495
      Left            =   4560
      MousePointer    =   99  'Custom
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   705
      Index           =   0
      Left            =   8760
      MousePointer    =   99  'Custom
      Top             =   6000
      Width           =   3090
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   1
      Left            =   3360
      MousePointer    =   99  'Custom
      Top             =   240
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   210
      Index           =   2
      Left            =   4560
      MousePointer    =   99  'Custom
      Top             =   480
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image FONDO 
      Height          =   9000
      Left            =   0
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Matías Fernando Pequeño
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Código Postal 1405

Option Explicit

Public Sub CargarLst()

Dim i As Integer

lst_servers.Clear

If ServersRecibidos Then
    Call WriteVar(App.Path & "\init\sinfo.dat", "INIT", "Cant", UBound(ServersLst))
    For i = 1 To UBound(ServersLst)
        Call WriteVar(App.Path & "\init\sinfo.dat", "S" & i, "Desc", ServersLst(i).desc)
        Call WriteVar(App.Path & "\init\sinfo.dat", "S" & i, "IP", ServersLst(i).Ip)
        Call WriteVar(App.Path & "\init\sinfo.dat", "S" & i, "PJ", Str(ServersLst(i).Puerto))
        Call WriteVar(App.Path & "\init\sinfo.dat", "S" & i, "P2", Str(ServersLst(i).PassRecPort))
        lst_servers.AddItem ServersLst(i).Ip & ":" & ServersLst(i).Puerto & " - Desc:" & ServersLst(i).desc
    Next i
End If

End Sub

Private Sub Command1_Click()
CurServer = 0
IPdelServidor = IPTxt
PuertoDelServidor = PortTxt
End Sub


Private Sub Command2_Click()

frmMain.Inet1.URL = "http://ao.alkon.com.ar/admin/iplist2.txt"
RawServersList = frmMain.Inet1.OpenURL


If RawServersList = "" Then
    ServersRecibidos = False
    Call MsgBox("No se pudo cargar la lista de servidores")
    ReDim ServersLst(1)
    Exit Sub
Else
    ServersRecibidos = True
End If

Call InitServersList(RawServersList)
Call CargarLst

End Sub





Private Sub Form_Activate()
'On Error Resume Next

If ServersRecibidos Then
    If CurServer <> 0 Then
        IPTxt = ServersLst(1).Ip
        PortTxt = ServersLst(1).Puerto
    Else
        IPTxt = IPdelServidor
        PortTxt = PuertoDelServidor
    End If
    
    Call CargarLst
Else
    lst_servers.Clear
End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
        frmCargando.Show
        frmCargando.Refresh
        AddtoRichTextBox frmCargando.status, "Cerrando Argentum Online.", 0, 0, 0, 1, 0, 1
        
        Call SaveGameini
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        
        AddtoRichTextBox frmCargando.status, "Liberando recursos..."
        Call WriteVar(App.Path & "\init\version.dat", "VERSION", "Graficos", "0")
        frmCargando.Refresh
        LiberarObjetosDX
        AddtoRichTextBox frmCargando.status, "Hecho", 0, 0, 0, 1, 0, 1
        AddtoRichTextBox frmCargando.status, "¡¡Gracias por jugar Argentum Online!!", 0, 0, 0, 1, 0, 1
        frmCargando.Refresh
        Call UnloadAllForms
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Make Server IP and Port box visible
If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    
    'Port
    PortTxt.Visible = False
    'Label4.Visible = True
    
    'Server IP
    PortTxt.Text = "7666"
    IPTxt.Text = "192.168.0.2"
    IPTxt.Visible = False
    'Label5.Visible = True
    
    KeyCode = 0
    Exit Sub
End If

End Sub

Private Sub Form_Load()
    '[CODE 002]:MatuX
    EngineRun = False
    '[END]
    
 Dim j
 For Each j In Image1()
    j.Tag = "0"
 Next
 PortTxt.Text = Config_Inicio.Puerto
 
 FONDO.Picture = LoadPicture(App.Path & "\Graficos\Conectar.jpg")
Timer1.Enabled = False

 '[CODE]:MatuX
 '
 '  El código para mostrar la versión se genera acá para
 ' evitar que por X razones luego desaparezca, como suele
 ' pasar a veces :)
    version.Caption = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
 '[END]'

End Sub



Private Sub Image1_Click(index As Integer)

CurServer = 0
'IPdelServidor = IPTxt
PuertoDelServidor = PortTxt
IPdelServidor = "127.0.0.1" 'testeando, NicoNZ

Call Audio.PlayWave(SND_CLICK)

Select Case index
    Case 0
        
        If Musica Then
            Call Audio.PlayMIDI("7.mid")
        End If
        
        
        
        'frmCrearPersonaje.Show vbModal
        EstadoLogin = Dados
#If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
        End If
        frmMain.Socket1.HostName = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
#Else
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
        End If
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
        Me.MousePointer = 11

        
    Case 1
    
        frmOldPersonaje.Show vbModal
        
    Case 2
        On Error GoTo errH
        Call Shell(App.Path & "\RECUPERAR.EXE", vbNormalFocus)

End Select
Exit Sub

errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "Argentum Online")
End Sub

Private Sub Image2_Click()
Call Audio.PlayWave(SND_CLICK)

#If UsarWrench = 1 Then
            If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
    #Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
    #End If
            If frmConnect.MousePointer = 11 Then
                Exit Sub
            End If
           
           
            UserName = NombreTXT.Text
        Dim aux As String
        aux = PasswordTxt.Text
#If SeguridadAlkon Then
        UserPassword = md5.GetMD5String(aux)
        Call md5.MD5Reset
#Else
        UserPassword = aux
#End If
            If CheckUserData(False) = True Then
                'SendNewChar = False
                EstadoLogin = Normal
                Me.MousePointer = 11
    #If UsarWrench = 1 Then
                frmMain.Socket1.HostName = CurServerIp
                frmMain.Socket1.RemotePort = CurServerPort
                frmMain.Socket1.Connect
    #Else
                If frmMain.Winsock1.State <> sckClosed Then _
                    frmMain.Winsock1.Close
                frmMain.Winsock1.Connect CurServerIp, CurServerPort
    #End If
            End If
End Sub

Private Sub imgGetPass_Click()
On Error GoTo errH

    Call Audio.PlayWave(SND_CLICK)
    Call Shell(App.Path & "\RECUPERAR.EXE", vbNormalFocus)
    'Call frmRecuperar.Show(vbModal, frmConnect)
    Exit Sub
errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "Argentum Online")
End Sub

Private Sub imgServArgentina_Click()
    Call Audio.PlayWave(SND_CLICK)
    IPTxt.Text = IPdelServidor
    PortTxt.Text = PuertoDelServidor
End Sub

Private Sub imgServEspana_Click()
    Call Audio.PlayWave(SND_CLICK)
    IPTxt.Text = "62.42.193.233"
    PortTxt.Text = "7666"
End Sub



Private Sub lst_servers_Click()
If ServersRecibidos Then
    CurServer = lst_servers.listIndex + 1
    IPTxt = ServersLst(CurServer).Ip
    PortTxt = ServersLst(CurServer).Puerto
End If

End Sub

Private Sub PasswordTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Image2_Click
    End If
End Sub

Private Sub Timer1_Timer()
frmConnect.Label1.Caption = " "
frmConnect.Timer1.Enabled = False
frmConnect.MousePointer = 1
End Sub
