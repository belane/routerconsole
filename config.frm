VERSION 5.00
Begin VB.Form config 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuracion "
   ClientHeight    =   4890
   ClientLeft      =   5910
   ClientTop       =   5175
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton sal 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2200
      TabIndex        =   12
      Top             =   4300
      Width           =   1215
   End
   Begin VB.TextBox pemuletcp 
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
      Left            =   360
      TabIndex        =   4
      Text            =   "4662;4242;"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox ipdft 
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
      Left            =   360
      TabIndex        =   10
      Text            =   "192.168.0.2"
      ToolTipText     =   "no lo voy a hacer yo por vosotros"
      Top             =   4430
      Width           =   1455
   End
   Begin VB.Frame fipdft 
      Caption         =   " IP Local "
      Height          =   600
      Left            =   90
      TabIndex        =   18
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton configreg 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   3480
      TabIndex        =   11
      Top             =   4300
      Width           =   1190
   End
   Begin VB.TextBox pserverudp 
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
      Left            =   2040
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox pservertcp 
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
      Left            =   360
      TabIndex        =   8
      Text            =   "80;"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox pnetbiosudp 
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
      Left            =   2040
      TabIndex        =   7
      Text            =   "137;138;"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox pnetbiostcp 
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
      Left            =   360
      TabIndex        =   6
      Text            =   "135;139;"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox pemuleudp 
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
      Left            =   2040
      TabIndex        =   5
      Text            =   "4672;"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame fparametros 
      Caption         =   " Parametros "
      Height          =   1455
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   4575
      Begin VB.CheckBox configmin 
         Caption         =   " Iniciar Minimizado."
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox configauto 
         Caption         =   " Conectar automaticamente al Iniciar."
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   3015
      End
      Begin VB.CheckBox configsave 
         Caption         =   " Guardar datos actuales de conexion con el Router.       (Ip, Puerto, Pass)"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame fpuertos 
      Caption         =   " Puertos "
      Height          =   2700
      Left            =   90
      TabIndex        =   13
      Top             =   1490
      Width           =   4575
      Begin VB.Label Label11 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "UDP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3240
         TabIndex        =   24
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "TCP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1560
         TabIndex        =   23
         Top             =   1680
         Width           =   330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "UDP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3240
         TabIndex        =   22
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "TCP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1560
         TabIndex        =   21
         Top             =   1080
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Netbios"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "HTTP"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "UDP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3240
         TabIndex        =   16
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "TCP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1560
         TabIndex        =   15
         Top             =   480
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Server"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "TCP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3720
      TabIndex        =   19
      Top             =   3000
      Width           =   330
   End
End
Attribute VB_Name = "config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub configreg_Click()
    If configsave.Value = 1 Then
        If principal.iprouter.Text <> "" Then Call aregistro("\iprouter", principal.iprouter.Text)
        If principal.puertorouter.Text <> "" Then Call aregistro("\puertorouter", principal.puertorouter.Text)
        If principal.passrouter.Text <> "nOcillero xD" And principal.passrouter.Text <> "" Then Call aregistro("\passrouter", encrip(principal.passrouter.Text))
        configsave.Value = 0
    End If
    If configauto.Value = 1 Then Call aregistro("\autoconn", "1")
    If configauto.Value = 0 Then Call aregistro("\autoconn", "0")
    If configmin.Value = 1 Then Call aregistro("\minconn", "1")
    If configmin.Value = 0 Then Call aregistro("\minconn", "0")
    If pemuletcp.Text <> "" Then Call aregistro("\pemuletcp", pemuletcp.Text)
    If pemuleudp.Text <> "" Then Call aregistro("\pemuleudp", pemuleudp.Text)
    If pnetbiostcp.Text <> "" Then Call aregistro("\pnetbiostcp", pnetbiostcp.Text)
    If pnetbiosudp.Text <> "" Then Call aregistro("\pnetbiosudp", pnetbiosudp.Text)
    If pservertcp.Text <> "" Then Call aregistro("\pservertcp", pservertcp.Text)
    If pserverudp.Text <> "" Then Call aregistro("\pserverudp", pserverudp.Text)
    If ipdft.Text <> "" Then Call aregistro("\ipdft", ipdft.Text)
    config.Hide
End Sub

Private Sub sal_Click()
    config.Hide
End Sub
