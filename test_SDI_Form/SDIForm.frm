VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form SDIForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SDIForm"
   ClientHeight    =   5520
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7725
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   960
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   240
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   8421504
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   74
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":0000
            Key             =   "check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":015A
            Key             =   "checkxp"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":02B4
            Key             =   "button"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":040E
            Key             =   "standard"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":0568
            Key             =   "xp"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":06C2
            Key             =   "open"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":120C
            Key             =   "new"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":1366
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":14C0
            Key             =   "print"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":161A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":1BB4
            Key             =   "saveas"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":1D0E
            Key             =   "save"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":22AA
            Key             =   "close"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":2404
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":255E
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":26B8
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":2812
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":296C
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":2AC6
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":2C20
            Key             =   "format"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":2D7A
            Key             =   "fontcolor"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":2ED4
            Key             =   "fontfolder"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":302E
            Key             =   "front"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":3188
            Key             =   "back"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":32E2
            Key             =   "bar3d"
            Object.Tag             =   "bar3d"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":343C
            Key             =   "line3d"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":3596
            Key             =   "pie3d"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":36F0
            Key             =   "help"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":384A
            Key             =   "helpfind"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":39A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":3AFE
            Key             =   "rect"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":3C58
            Key             =   "arc"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":3DB2
            Key             =   "draw"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":3F0C
            Key             =   "elipse"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":4066
            Key             =   "freeform"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":41C0
            Key             =   "line"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":431A
            Key             =   "wincascade"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":4474
            Key             =   "winhoriz"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":45CE
            Key             =   "winvert"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":4728
            Key             =   "winicons"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":4882
            Key             =   "webhome"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":49DC
            Key             =   "weblink"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":4B36
            Key             =   "webopen"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":4C90
            Key             =   "websearch"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":4DEA
            Key             =   "webend"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":4F44
            Key             =   "mail"
            Object.Tag             =   "mail"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":509E
            Key             =   "color"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":51F8
            Key             =   "xpbar"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":5352
            Key             =   "xpback"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":54AC
            Key             =   "xpfill"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":5606
            Key             =   "xpborder"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":5760
            Key             =   "xpfill2"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":58BA
            Key             =   "menubar"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":5A14
            Key             =   "menuimg"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":5B6E
            Key             =   "separator"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":5CC8
            Key             =   "menu"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":5E22
            Key             =   "menu2"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":5F7C
            Key             =   "info"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":60D6
            Key             =   "uomo"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":61F2
            Key             =   "clipboard"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":634C
            Key             =   "vbform"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":64A6
            Key             =   "vbchild"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":6600
            Key             =   "vbmdi"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":675A
            Key             =   "plus"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":68B4
            Key             =   "meno"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":6A0E
            Key             =   "menuselector"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":6B68
            Key             =   "menuselector1"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":6CC2
            Key             =   "menuselector2"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":6E1C
            Key             =   "smallicons"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":6F76
            Key             =   "largeicons"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":70D0
            Key             =   "listicons"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":722A
            Key             =   "detailicons"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":7384
            Key             =   "toolbar"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDIForm.frx":74DE
            Key             =   "prop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5205
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "gibra"
            TextSave        =   "gibra"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   2280
      Picture         =   "SDIForm.frx":7638
      Top             =   3420
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Menu mnuBar 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "&Nuovo"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Apri"
         Index           =   1
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFile 
         Caption         =   "C&hiudi"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-separatore"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Salva"
         Enabled         =   0   'False
         Index           =   4
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Stam&pa..."
         Index           =   6
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Es&ci"
         Index           =   8
      End
   End
   Begin VB.Menu mnuBar 
      Caption         =   "&Opzioni"
      Index           =   1
      Begin VB.Menu mnuOpzioni 
         Caption         =   "&Preferenze..."
         Index           =   0
      End
      Begin VB.Menu mnuOpzioni 
         Caption         =   "&Registra posizione finestra"
         Index           =   1
      End
      Begin VB.Menu mnuOpzioni 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuOpzioni 
         Caption         =   "&Menu Designer"
         Index           =   3
      End
   End
   Begin VB.Menu mnuBar 
      Caption         =   "&Visualizza"
      Index           =   2
      Begin VB.Menu mnuView 
         Caption         =   "&Barra degli strumenti"
         Index           =   0
      End
      Begin VB.Menu mnuView 
         Caption         =   "Barra di &stato"
         Index           =   1
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuView 
         Caption         =   "Icone g&randi"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuView 
         Caption         =   "Icone pi&ccole"
         Index           =   4
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Elenco"
         Index           =   5
      End
      Begin VB.Menu mnuView 
         Caption         =   "De&ttagli"
         Index           =   6
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Disponi icone"
         Index           =   8
         Begin VB.Menu mnuDisponiIcone 
            Caption         =   "Per &nome"
            Index           =   0
         End
         Begin VB.Menu mnuDisponiIcone 
            Caption         =   "Per &tipo"
            Index           =   1
         End
         Begin VB.Menu mnuDisponiIcone 
            Caption         =   "Per di&mensione"
            Index           =   2
         End
         Begin VB.Menu mnuDisponiIcone 
            Caption         =   "Per &data"
            Index           =   3
         End
      End
      Begin VB.Menu mnuView 
         Caption         =   "Alli&nea icone"
         Index           =   9
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Opzioni..."
         Index           =   11
      End
   End
   Begin VB.Menu mnuBar 
      Caption         =   "&?"
      Index           =   3
      Begin VB.Menu mnuHelp 
         Caption         =   "&Sommario"
         Index           =   0
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Cerca file della guida..."
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Informazioni su ..."
         Index           =   3
      End
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "Hidden"
      Begin VB.Menu mnuPopup 
         Caption         =   "Questo menu"
         Index           =   0
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "è nascosto"
         Index           =   1
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "e l'ho messo"
         Index           =   2
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "sulla barra"
         Index           =   3
      End
   End
End
Attribute VB_Name = "SDIForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Eventi As CEvents
Attribute Eventi.VB_VarHelpID = -1
Dim ApplicareClasse As Boolean


Private Sub Eventi_MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal Enabled As Boolean)
  If Enabled Then
    sb.Panels(1).Text = MenuHelp$
  Else
    sb.Panels(1).Text = ""
  End If
End Sub


Private Sub Form_Load()
  
  On Error GoTo Err_load
  ' nascondo il menu opzioni
  mnuHidden.Visible = False

  
  ApplicareClasse = True  ' Flag per non utilizzare la classe MenuEx

  If ApplicareClasse Then
    mnuFile(0).Caption = mnuFile(0).Caption & "|@Crea nuovo documento|#new"
    mnuFile(1).Caption = mnuFile(1).Caption & "|@Apre Form2, fai clic per provare PopupMenu|#open"
    mnuFile(2).Caption = mnuFile(2).Caption & "|@Chiude il documento aperto|#close"
    mnuFile(4).Caption = mnuFile(4).Caption & "|@Salva il documento|#save"
    mnuFile(6).Caption = mnuFile(6).Caption & "|@Stampa il documento|#print"
    mnuFile(8).Caption = mnuFile(8).Caption & "|@Chiude il programma|#exit"
    
    
    
    mnuHelp(0).Caption = mnuHelp(0).Caption & "|@Apre la guida in linea|#help"
    mnuHelp(1).Caption = mnuHelp(1).Caption & "|@Cerca un argomento nella guida|#helpfind"
    mnuHelp(3).Caption = mnuHelp(3).Caption & "|@Informazioni sul programma|#info"
    
    mnuView(0).Caption = mnuView(0).Caption & "|@Mostra/nasconde la barra degli strumenti|#toolbar"
    mnuView(1).Caption = mnuView(1).Caption & "|@Mostra/nasconde la barra di stato"  '|#statbar"
    mnuView(3).Caption = mnuView(3).Caption & "|@Visualizza icone grandi|#largeicons"
    mnuView(4).Caption = mnuView(4).Caption & "|@Visualizza icone piccole|#smallicons"
    mnuView(5).Caption = mnuView(5).Caption & "|@Visualizza elenco|#listicons"
    mnuView(6).Caption = mnuView(6).Caption & "|@Visualizza dettagli|#detailicons"
    mnuView(11).Caption = mnuView(11).Caption & "|@Apre la finestra delle opzioni|#prop"
    
    mnuDisponiIcone(0).Caption = mnuDisponiIcone(0).Caption & "|@Clic per selezionare l'opzione"
    mnuDisponiIcone(1).Caption = mnuDisponiIcone(1).Caption & "|@Clic per selezionare l'opzione"
    mnuDisponiIcone(2).Caption = mnuDisponiIcone(2).Caption & "|@Clic per selezionare l'opzione"
    mnuDisponiIcone(3).Caption = mnuDisponiIcone(3).Caption & "|@Clic per selezionare l'opzione"
    
    mnuOpzioni(0).Caption = mnuOpzioni(0).Caption & "|@Personalizza il programma|#prop"
    mnuOpzioni(1).Caption = mnuOpzioni(1).Caption & "|@Memorizza posizione e dimensione della finestra"
    mnuOpzioni(3).Caption = mnuOpzioni(3).Caption & "|@Personalizza i colori|#color"
    
    mnuPopup(0).Caption = mnuPopup(0).Caption & "|@Questo menu è l'ultimo menu a destra sulla barra|#menu2"
    mnuPopup(1).Caption = mnuPopup(1).Caption & "|@Anche se sposti il mouse sulla barra dei menu, non mi vedi|#menu2"
    mnuPopup(2).Caption = mnuPopup(2).Caption & "|@Che dici Mark, può andare così?|#menu2"
    mnuPopup(3).Caption = mnuPopup(3).Caption & "|@Devi solo renderlo invisibile.|#menu2"
    
    Set Eventi = New CEvents
    Set objMenuEx = New cMenuEx
    Call objMenuEx.Install(SDIForm.hWnd, App.Path & "\" & Me.Name, iml, 0, Eventi)
    
    
  End If
  
Exit Sub

Err_load:
  Debug.Print Err.Number & ": " & Err.Description
  Resume Next
  
End Sub





Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button And 2 Then
    ' se è attiva la classe MenuEx
    If objMenuEx Is Nothing Then
      
      If Shift And vbCtrlMask Then
        ' questo menu è nel menu Modifica
        PopupMenu mnuBar(2)
      Else
        ' questo menu è sulla barra
        PopupMenu mnuHidden
      End If
      
    Else
    
      If Shift And vbCtrlMask Then
        ' questo menu è nel menu Modifica
        Call objMenuEx.PopupMenu(Me, Me.mnuBar(2))
      Else
        ' questo menu è sulla barra
        'mnuPopup(0).Caption = "Questa voce l'ho modificata a runtime per vedere se si verificano errori|@Nuova Descrizione|#new"
        Call objMenuEx.PopupMenu(Me, Me.mnuHidden)
      End If
    End If
    
  End If

End Sub

Private Sub Form_Resize()
  CurrentY = 0
  CurrentX = 0
  Font.Bold = False
  Print ""
'  Print "   Ciao Mark,"
  Print ""
  Font.Bold = True
  Print "   Per aprire un menu popup da questo form:"
  Font.Bold = False
  Print "     - Clic Destro apre il menu nascosto sulla barra, a destra dopo ?"
  Print "     - Clic Destro insieme a CTRL apre il menu nascosto nel menu Modifica"
  Print ""
  Font.Bold = True
  Print "   Per aprire un menu popup da un altro form:"
  Font.Bold = False
  Print "     Seleziona File Apri per aprire Form2"
  Print "     e ripeti come sopra."
'  Print ""
'  Font.Bold = True
'  Print "   Come vedi, in nessuno dei casi si notano i due menu nascosti."
'  Print "   Va bene così?"

End Sub


Private Sub Form_Unload(Cancel As Integer)

  If ApplicareClasse Then
    Call objMenuEx.Uninstall(Me.hWnd, iml, Eventi)
    Set Eventi = Nothing
  End If

End Sub



Private Sub Image1_Click(Index As Integer)

End Sub

Private Sub mnuDisponiIcone_Click(Index As Integer)
  Dim i As Integer
  For i = mnuDisponiIcone.LBound To mnuDisponiIcone.UBound
    mnuDisponiIcone(i).Checked = False
  Next i
  mnuDisponiIcone(Index).Checked = True
End Sub

Private Sub mnuFile_Click(Index As Integer)
  Select Case Index
    Case 0  ' attiva/disattivo la bitmap Open
      'mnuFile(1).Enabled = Not mnuFile(1).Enabled
    Case 1
      Form2.Show vbModal, Me
    Case 8
      Unload Me
  End Select
  
End Sub


Private Sub mnuHelp_Click(Index As Integer)
Dim nRet As Integer
On Error Resume Next

Select Case Index
  Case 0
    SendKeys "{F1}"
  Case 1
  
  Case 3
    MsgBox "Informazioni su..." & vbCr & vbCr & _
          "MenuExtended (MenuEx.dll)" & vbCr & _
          "   di Giorgio Brausi"
    
  End Select
  
End Sub


Private Sub mnuOpzioni_Click(Index As Integer)
  Select Case Index
    Case 0  '
      MsgBox "Hai scelto l'opzione:" & vbCr & "Preferenze"
    Case 1
      mnuOpzioni(Index).Checked = Not mnuOpzioni(Index).Checked
      MsgBox "Hai scelto l'opzione:" & vbCr & "Registra posizione finestra" & vbCr & "(" & mnuOpzioni(Index).Checked & ")"
    Case 3
        objMenuEx.MenuDesigner Me.hWnd
  End Select
End Sub

Private Sub mnuView_Click(Index As Integer)
Dim i As Integer
  Select Case Index
    Case 0, 1, 9
      mnuView(Index).Checked = Not mnuView(Index).Checked
    Case 3, 4, 5, 6
      For i = 3 To 6
        mnuView(i).Checked = False
      Next i
      mnuView(Index).Checked = True
  End Select
End Sub


