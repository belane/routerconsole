VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form principal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Remote Management Console"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fconfig 
      Caption         =   " Configuracion "
      Height          =   810
      Left            =   120
      TabIndex        =   32
      Top             =   5090
      Width           =   1935
      Begin VB.CommandButton acLimpiarC 
         Caption         =   "Limpiar Consola"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   470
         Width           =   1695
      End
      Begin VB.CommandButton acConfig 
         Caption         =   "Config Consola"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   230
         Width           =   1695
      End
   End
   Begin VB.Timer intruc 
      Left            =   2040
      Top             =   1680
   End
   Begin VB.CommandButton cerrar 
      Appearance      =   0  'Flat
      Caption         =   "Desconectar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1235
      Width           =   1695
   End
   Begin VB.CommandButton inicializar 
      Appearance      =   0  'Flat
      Caption         =   "Conectar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   990
      Width           =   1695
   End
   Begin VB.TextBox exec 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   5490
      Width           =   4515
   End
   Begin VB.ComboBox execrap 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "main.frx":1E0A3
      Left            =   7380
      List            =   "main.frx":1E170
      OLEDragMode     =   1  'Automatic
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5490
      Width           =   1455
   End
   Begin VB.CommandButton paya 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "<<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6840
      MaskColor       =   &H00808080&
      TabIndex        =   1
      Top             =   5484
      Width           =   495
   End
   Begin VB.Frame facciones 
      Caption         =   " Acciones Rapidas "
      Enabled         =   0   'False
      Height          =   3440
      Left            =   120
      TabIndex        =   11
      Top             =   1640
      Width           =   1935
      Begin VB.CommandButton acReboot 
         Caption         =   "Reiniciar Router"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3060
         Width           =   1695
      End
      Begin VB.CommandButton acCfg 
         Caption         =   "Mostrar Config"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2820
         Width           =   1695
      End
      Begin VB.CommandButton swebOFF 
         Caption         =   "OFF "
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton swebON 
         Caption         =   "  HTTP   ON"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton NetbibOFF 
         Caption         =   "OFF "
         Height          =   255
         Left            =   1320
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton NetbiON 
         Caption         =   "Netbios  ON"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton acStatTU 
         Caption         =   "TCP/UDP Stats"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2580
         Width           =   1695
      End
      Begin VB.CommandButton eMOFF 
         Caption         =   "OFF "
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton eMON 
         Caption         =   "Server   ON"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton acSRouter 
         Caption         =   "Estado Router"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1695
      End
      Begin VB.CommandButton SFcfg 
         Caption         =   "Ver Filtros"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1695
      End
      Begin VB.CommandButton SetFOff 
         Caption         =   "OFF"
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1020
         Width           =   495
      End
      Begin VB.CommandButton SetFOn 
         Caption         =   "ON"
         Height          =   255
         Left            =   840
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1020
         Width           =   495
      End
      Begin VB.CommandButton SFilter 
         Caption         =   "Filtros"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1020
         Width           =   735
      End
      Begin VB.CommandButton SNaptS 
         Caption         =   "Servidores NAPT"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton acLNaptM 
         Caption         =   "Limpiar NAPT Act"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton acSNaptM 
         Caption         =   "Conexiones NAPT"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fconexion 
      Caption         =   " Conexion "
      Height          =   1500
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1935
      Begin VB.TextBox passrouter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "-"
         TabIndex        =   6
         Text            =   "nOcillero xD"
         Top             =   550
         Width           =   1695
      End
      Begin VB.TextBox puertorouter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1440
         TabIndex        =   5
         Text            =   "23"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox iprouter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Text            =   "192.168.0.1"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label puntos 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1350
         TabIndex        =   10
         Top             =   240
         Width           =   75
      End
   End
   Begin VB.CommandButton comando 
      Appearance      =   0  'Flat
      Caption         =   "Ejecutar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8880
      TabIndex        =   3
      Top             =   5484
      Width           =   1095
   End
   Begin VB.Frame fconsola 
      Caption         =   " Consola "
      Height          =   5775
      Left            =   2160
      TabIndex        =   12
      Top             =   120
      Width           =   7935
      Begin VB.PictureBox IconoBandeja 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         Picture         =   "main.frx":1E3BC
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin RichTextLib.RichTextBox recibido 
         Height          =   5055
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   8916
         _Version        =   393217
         BackColor       =   -2147483633
         BorderStyle     =   0
         Enabled         =   -1  'True
         MousePointer    =   1
         Appearance      =   0
         TextRTF         =   $"main.frx":3C45F
      End
   End
   Begin MSWinsockLib.Winsock conect 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu MenuSystray 
      Caption         =   "MenuSystray"
      Visible         =   0   'False
      Begin VB.Menu msConectar 
         Caption         =   "Conectar"
         Index           =   1
      End
      Begin VB.Menu msDesconectar 
         Caption         =   "Desconectar"
         Enabled         =   0   'False
      End
      Begin VB.Menu linea1 
         Caption         =   "-"
      End
      Begin VB.Menu msFiltros 
         Caption         =   "Filtros"
         Begin VB.Menu msFiltrosOn 
            Caption         =   "ON"
            Enabled         =   0   'False
         End
         Begin VB.Menu msFiltrosOff 
            Caption         =   "OFF"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mseMule 
         Caption         =   "eMule"
         Begin VB.Menu mseMuleON 
            Caption         =   "ON"
            Enabled         =   0   'False
         End
         Begin VB.Menu mseMuleOFF 
            Caption         =   "OFF"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu msconexiones 
         Caption         =   "Conexiones"
         Enabled         =   0   'False
      End
      Begin VB.Menu linea2 
         Caption         =   "-"
      End
      Begin VB.Menu msAbrir 
         Caption         =   "Abrir"
      End
      Begin VB.Menu msSalir 
         Caption         =   "Cerrar"
      End
   End
End
Attribute VB_Name = "principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ip_router As String
Public puerto_router As Integer
Public pass As String
Public ultimo_exec As String
Public lukpass As Boolean
Dim ultimos(20) As String
Dim posulti As Integer
Dim primera As Boolean
Dim tim As Integer

Private Sub acConfig_Click()
    config.Show 1
End Sub

Private Sub cerrar_Click()
    inicializar.Enabled = True
    conect.Close
    cerrar.Enabled = False
    comando.Enabled = False
    exec.Enabled = False
    execrap.Enabled = False
    inicializar.Caption = "Conectar"
    msDesconectar.Enabled = False
    msconexiones.Enabled = False
    msFiltrosOff.Enabled = False
    msFiltrosOn.Enabled = False
    mseMuleON.Enabled = False
    mseMuleOFF.Enabled = False
    facciones.Enabled = False
    paya.Enabled = False
    msConectar.Item(1).Caption = "Conectar"
    ShellTrayModifyTip BALOON_ICO_INFO, "Consola Router", "Desconectado"
End Sub

Private Sub comando_Click()
    conect.SendData exec.Text & vbCrLf
    acumula_ultimos (exec.Text)
    exec.Text = ""
    exec.SetFocus
    posulti = 0
End Sub

Private Sub conect_Close()
    recibido = recibido & vbCrLf & "  *** DESCONECTADO ***" & vbCrLf
    recibido.SelStart = Len(recibido.Text)
    Call cerrar_Click
End Sub

Private Sub conect_Connect()
    inicializar.Enabled = True
    cerrar.Enabled = True
    msDesconectar.Enabled = True
    comando.Enabled = True
    exec.Enabled = True
    execrap.Enabled = True
    inicializar.Caption = "Reconectar"
    msConectar.Item(1).Caption = "Reconectar"
    msconexiones.Enabled = Enabled
    msFiltrosOff.Enabled = Enabled
    msFiltrosOn.Enabled = Enabled
    mseMuleON.Enabled = Enabled
    mseMuleOFF.Enabled = Enabled
    facciones.Enabled = True
    paya.Enabled = True
    ShellTrayModifyTip BALOON_ICO_INFO, "Consola Router", "Conectado"
    'If principal.WindowState <> 1 Then exec.SetFocus
End Sub

Private Sub conect_DataArrival(ByVal bytesTotal As Long)
    Dim datos As String
    conect.GetData datos
    If InStr(datos, "Press any key") > 0 Then conect.SendData " "
    datos = Replace(Replace(Replace(datos, Chr$(13), ""), Chr$(10), vbCrLf), "ÿû", "")
    If InStr(datos, "Password:") > 0 And lukpass = True Then
        Call MsgBox("Error de autentificacion", vbInformation, "Password Incorrecta")
        lukpass = False
        cerrar_Click
        Exit Sub
    End If
    If InStr(datos, "Password:") > 0 And lukpass = False Then
        lukpass = True
        conect.SendData pass & vbCrLf
    End If
    recibido.Text = Replace(recibido.Text, "-- Press any key to continue --" & vbCrLf, "") & datos
    recibido.SelStart = Len(recibido.Text)
End Sub

Private Sub conect_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call MsgBox(vbCrLf & Description & vbCrLf & "Verifique los datos de conexion." & vbCrLf & "  [" & Number & "|" & Scode & "]", 64, "ConsolaSS: Error! ")
    conect.Close
    inicializar.Enabled = True
End Sub


Private Sub exec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
            exec.Text = ultimos(posulti)
            If posulti <> 19 And ultimos(posulti + 1) <> "" Then posulti = posulti + 1
    End If
    If KeyCode = 40 Then
            If posulti <> 0 Then posulti = posulti - 1
            exec.Text = ultimos(posulti)
    End If
End Sub
Private Sub exec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call comando_Click
End Sub

Private Sub execrap_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call comando_Click
    If KeyAscii = 32 Then exec.Text = exec.Text & execrap & " "
End Sub

Private Sub Form_Load()
    On Error GoTo fin
    If QueryRegBase("software\consolass\minconn") = 1 Then
        principal.Hide
        ShellTrayAdd ("Router")
        config.configmin.Value = 1
    End If
    If QueryRegBase("software\consolass\iprouter") <> "" Then principal.iprouter.Text = QueryRegBase("software\consolass\iprouter")
    If QueryRegBase("software\consolass\puertorouter") <> "" Then principal.puertorouter.Text = QueryRegBase("software\consolass\puertorouter")
    If QueryRegBase("software\consolass\passrouter") <> "" Then pass = dencrip(QueryRegBase("software\consolass\passrouter"))
    If QueryRegBase("software\consolass\autoconn") = 1 Then
        config.configauto.Value = 1
        inicializar_Click
    End If
    If QueryRegBase("software\consolass\pemuletcp") <> "" Then config.pemuletcp.Text = QueryRegBase("software\consolass\pemuletcp")
    If QueryRegBase("software\consolass\pemuleudp") <> "" Then config.pemuleudp.Text = QueryRegBase("software\consolass\pemuleudp")
    If QueryRegBase("software\consolass\pnetbiostcp") <> "" Then config.pnetbiostcp.Text = QueryRegBase("software\consolass\pnetbiostcp")
    If QueryRegBase("software\consolass\pnetbiosudp") <> "" Then config.pnetbiosudp.Text = QueryRegBase("software\consolass\pnetbiosudp")
    If QueryRegBase("software\consolass\pservertcp") <> "" Then config.pservertcp.Text = QueryRegBase("software\consolass\pservertcp")
    If QueryRegBase("software\consolass\pserverudp") <> "" Then config.pserverudp.Text = QueryRegBase("software\consolass\pserverudp")
    If QueryRegBase("software\consolass\ipdft") <> "" Then config.ipdft.Text = QueryRegBase("software\consolass\ipdft")
fin:
End Sub

Private Sub Form_Resize()
    If principal.WindowState = 1 Then
        ShellTrayAdd ("Router")
        If primera = False Then
            If conect.State = 7 Then
                entonces = "Conectado"
            Else
                entonces = "Desconectado"
            End If
            ShellTrayModifyTip BALOON_ICO_INFO, "Consola Router", "Consola Activa [" & entonces & "]"
        End If
        primera = True
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    conect.Close
    Unload config
    Call ShellTrayRemove
End Sub

Private Sub IconoBandeja_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim butt
butt = X / Screen.TwipsPerPixelX
    Select Case butt
            Case WM_LBUTTONUP
                PopupMenu MenuSystray, 0
            Case WM_LBUTTONDBLCLK
                principal.WindowState = 0
                Call ShellTrayRemove
                principal.Show
            Case WM_RBUTTONUP
                PopupMenu MenuSystray, 0
            Case Else
                Exit Sub
    End Select
End Sub

Private Sub inicializar_Click()
    inicializar.Enabled = False
    If inicializar.Caption = "Reconectar" Then
        conect.SendData "exit"
        conect.Close
    End If
    ip_router = iprouter.Text
    puerto_router = puertorouter.Text
    If passrouter.Text <> "nOcillero xD" Then pass = passrouter.Text
    lukpass = False
    conect.Protocol = sckTCPProtocol
    conect.Connect ip_router, puerto_router
End Sub
Private Sub acumula_ultimos(ultimo_comando As String)
    For i = 20 To 1 Step -1
        ultimos(i) = ultimos(i - 1)
    Next
    ultimos(0) = ultimo_comando
End Sub

Private Sub intruc_Timer()
    Select Case tim
    Case 0
        If IntrucData(0) <> "" Then conect.SendData IntrucData(0) & vbCrLf
        IntrucData(0) = ""
        tim = tim + 1
    Case 1
        If IntrucData(1) <> "" Then conect.SendData IntrucData(1) & vbCrLf
        IntrucData(1) = ""
        tim = tim + 1
    Case 2
        If IntrucData(2) <> "" Then conect.SendData IntrucData(2) & vbCrLf
        IntrucData(2) = ""
        tim = tim + 1
    Case 3
        If IntrucData(3) <> "" Then conect.SendData IntrucData(3) & vbCrLf
        IntrucData(3) = ""
        tim = tim + 1
    Case 4
        If IntrucData(4) <> "" Then conect.SendData IntrucData(4) & vbCrLf
        IntrucData(0) = ""
        tim = 0
        intruc.Interval = 0
    Case Else
        tim = 0
        intruc.Interval = 0
    End Select
    DoEvents
End Sub

Private Sub msAbrir_Click()
    principal.WindowState = 0
    Call ShellTrayRemove
    principal.Show
End Sub

Private Sub msConectar_Click(Index As Integer)
    inicializar_Click
End Sub

Private Sub msDesconectar_Click()
    cerrar_Click
End Sub

Private Sub msconexiones_Click()
    recibido.Text = ""
    conect.SendData "Show naptmap" & vbCrLf
    msAbrir_Click
End Sub

Private Sub mseMuleOFF_Click()
    Call eMOFF_Click
End Sub

Private Sub mseMuleON_Click()
 Call eMON_Click
End Sub

Private Sub msFiltrosOff_Click()
    conect.SendData "set ipfilter disable" & vbCrLf
End Sub

Private Sub msFiltrosOn_Click()
    conect.SendData "set ipfilter enable" & vbCrLf
End Sub

Private Sub msSalir_Click()
    Call ShellTrayRemove
    conect.Close
    Unload principal
End Sub

Private Sub paya_Click()
    exec.Text = exec.Text & execrap & " "
End Sub

'  quick action buttons
Private Sub SetFOff_Click()
    conect.SendData "set ipfilter disable" & vbCrLf
End Sub

Private Sub SetFOn_Click()
    conect.SendData "set ipfilter enable" & vbCrLf
End Sub

Private Sub SFcfg_Click()
    conect.SendData "show ipfiltercfg" & vbCrLf
End Sub

Private Sub SFilter_Click()
    conect.SendData "show ipfilter" & vbCrLf & "show ipfiltercfg" & vbCrLf
End Sub

Private Sub SNaptS_Click()
    conect.SendData "Show naptserver" & vbCrLf
End Sub

Private Sub acSRouter_Click()
    conect.SendData "Show" & vbCrLf
End Sub
Private Sub acLimpiar_Click()
    recibido.Text = ""
End Sub

Private Sub acLimpiarC_Click()
    recibido.Text = ""
End Sub

Private Sub acLNaptM_Click()
    If MsgBox("Limpiar la Tabla NAPT de conexiones" & vbCrLf & "finalizara todas las conexiones actuales" & vbCrLf & "establecidas desde o hacia Internet," & vbCrLf & "¿Desea proceder de todas formas?", vbInformation + vbYesNo, "Atencion!") = 6 Then conect.SendData "clear naptmap" & vbCrLf
End Sub

Private Sub acSNaptM_Click()
    conect.SendData "show naptmap" & vbCrLf
End Sub
Private Sub eMON_Click()
    Call interp(config.pemuletcp.Text, config.pemuleudp.Text, "s")
    intruc.Interval = 250
End Sub
Private Sub eMOFF_Click()
    Call interp(config.pemuletcp.Text, config.pemuleudp.Text, "d")
    intruc.Interval = 250
End Sub
Private Sub NetbiON_Click()
    Call interp(config.pnetbiostcp.Text, config.pnetbiosudp.Text, "s")
    intruc.Interval = 250
End Sub
Private Sub NetbibOFF_Click()
    Call interp(config.pnetbiostcp.Text, config.pnetbiosudp.Text, "d")
    intruc.Interval = 250
End Sub
Private Sub swebON_Click()
    Call interp(config.pservertcp.Text, config.pserverudp.Text, "s")
    intruc.Interval = 250
End Sub
Private Sub swebOFF_Click()
    Call interp(config.pservertcp.Text, config.pserverudp.Text, "d")
    intruc.Interval = 250
End Sub
Private Sub acReboot_Click()
    IntrucData(0) = "reboot"
    IntrucData(1) = "y"
    intruc.Interval = 310
End Sub
Private Sub acStatTU_Click()
    IntrucData(0) = "show tcpstats"
    IntrucData(1) = "show udpstats"
    intruc.Interval = 310
End Sub
Private Sub acCfg_Click()
    conect.SendData "show cfg" & vbCrLf
End Sub


