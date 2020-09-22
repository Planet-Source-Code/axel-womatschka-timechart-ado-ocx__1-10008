VERSION 5.00
Begin VB.UserControl Urlaub 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "timechart.ctx":0000
   ScaleHeight     =   2370
   ScaleWidth      =   6990
   Begin VB.Frame Frame_ALL 
      BorderStyle     =   0  'Kein
      Height          =   2295
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame_Legende 
         BorderStyle     =   0  'Kein
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   450
         Width           =   6375
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "halber Tag"
            Height          =   225
            Left            =   4770
            TabIndex        =   16
            Top             =   30
            Width           =   885
         End
         Begin VB.Shape Shape_LEG_HALBER_TAG 
            FillColor       =   &H00C00000&
            FillStyle       =   0  'Ausgefüllt
            Height          =   195
            Left            =   4530
            Shape           =   3  'Kreis
            Top             =   60
            Width           =   165
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "voller Tag"
            Height          =   225
            Left            =   3540
            TabIndex        =   15
            Top             =   30
            Width           =   795
         End
         Begin VB.Shape Shape_LEG_EIN_TAG 
            FillColor       =   &H00C00000&
            FillStyle       =   0  'Ausgefüllt
            Height          =   195
            Left            =   3300
            Shape           =   3  'Kreis
            Top             =   60
            Width           =   165
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sonderurlaub"
            Height          =   225
            Left            =   1950
            TabIndex        =   14
            Top             =   30
            Width           =   1125
         End
         Begin VB.Label LBL_LEG_SONDERURLAUB 
            Alignment       =   2  'Zentriert
            Appearance      =   0  '2D
            BackColor       =   &H00FF0000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   1620
            TabIndex        =   13
            Top             =   30
            Width           =   285
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jahresurlaub"
            Height          =   225
            Left            =   360
            TabIndex        =   12
            Top             =   30
            Width           =   1095
         End
         Begin VB.Label LBL_LEG_JAHRESURLAUB 
            Alignment       =   2  'Zentriert
            Appearance      =   0  '2D
            BackColor       =   &H00FF0000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   30
            TabIndex        =   11
            Top             =   30
            Width           =   285
         End
      End
      Begin VB.Label LBL_U 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H00FF0000&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   2490
         MousePointer    =   14  'Pfeil und Fragezeichen
         TabIndex        =   9
         Top             =   1590
         Visible         =   0   'False
         Width           =   2985
      End
      Begin VB.Label LBL_HEADER 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00808000&
         Caption         =   "Header_KW"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   8
         Top             =   1230
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Line Line_KW 
         Index           =   0
         Visible         =   0   'False
         X1              =   2370
         X2              =   2370
         Y1              =   2280
         Y2              =   4380
      End
      Begin VB.Line Line_top 
         BorderWidth     =   2
         X1              =   90
         X2              =   3930
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   "Mitarbeiter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   90
         TabIndex        =   7
         Top             =   900
         Width           =   2220
      End
      Begin VB.Label LBL_NR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ma_nr"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   870
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label LBL_X 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   0
         Left            =   2280
         TabIndex        =   5
         Top             =   1230
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Line LINE_S 
         BorderWidth     =   2
         Index           =   0
         X1              =   2370
         X2              =   2370
         Y1              =   870
         Y2              =   4890
      End
      Begin VB.Line LINE_W 
         BorderWidth     =   2
         Index           =   0
         X1              =   90
         X2              =   3810
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Label LBL_Y 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   120
         MousePointer    =   14  'Pfeil und Fragezeichen
         TabIndex        =   4
         Top             =   1230
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Kalenderwoche"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2460
         TabIndex        =   3
         Top             =   900
         Width           =   2265
      End
      Begin VB.Shape Shape_EIN_TAG 
         FillColor       =   &H00C00000&
         FillStyle       =   0  'Ausgefüllt
         Height          =   195
         Index           =   0
         Left            =   4830
         Shape           =   3  'Kreis
         Top             =   900
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Shape Shape_HALBER_TAG 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Ausgefüllt
         Height          =   195
         Index           =   0
         Left            =   5130
         Shape           =   3  'Kreis
         Top             =   900
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label LBL_UEBERSCHRIFT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angezeigter Zeitraum:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   60
         Width           =   2085
      End
   End
   Begin VB.VScrollBar VScroll 
      CausesValidation=   0   'False
      Height          =   2265
      LargeChange     =   40
      Left            =   6600
      MousePointer    =   7  'Größenänderung N S
      SmallChange     =   20
      TabIndex        =   0
      Top             =   0
      Width           =   315
   End
End
Attribute VB_Name = "Urlaub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Standard Eigenschaftswerte:
Const m_def_ConnectionString = """"
'Const m_def_ConnectionString = 0
Const m_def_Filter_Bedingung = ""
Const m_def_Filter_Feldname = ""
Const m_def_ToolTip_Y = ""
Const m_def_ForeColor = 0
Const m_def_Legende = 1
'Const m_def_ForeColor = 0
Const m_def_Farbe_Urlaub = 0
Const m_def_Schriftfarbe_urlaub = 0
Const m_def_Farbe_HalberTag = 0
Const m_def_Farbe_einzelner_Tag = 0
Const m_def_Caption_Y_Achse = "Mitarbeiter"
Const m_def_TabFeld_ANZAHL_URLAUBSTAGE = "ANZAHL_URLAUBSTAGE"
Const m_def_Quartal = 1
Const m_def_Jahr = 2000
'Const m_def_Datenbankpfad = ""
Const m_def_Tabelle_DATUM = "URLAUB"
Const m_def_Tabelle_NAMEN = "MITARBEITER3"
Const m_def_TabFeld_Datum_von = "URLAUB_VON"
Const m_def_TabFeld_Datum_bis = "URLAUB_BIS"
Const m_def_TabFeld_Sonderurlaub = "SONDERURLAUB"
Const m_def_TabFeld_NAMEN = "KUERZEL"
Const m_def_TabFeld_NUMMER = "MA_NR"
Const m_def_Farbe_Sonderurlaub = &H80&
Const m_def_BorderStyle = 0
Const m_def_Enabled = True
Const m_def_Nullpunkt_x = 0
'Eigenschaft-Variablen:
Dim m_ConnectionString As String
'Dim m_ConnectionString As Variant
Dim m_Filter_Bedingung As String
Dim m_Filter_Feldname As String
Dim m_ToolTip_Y As String
Dim m_ForeColor As OLE_COLOR
Dim m_Legende As Boolean
'Dim m_ForeColor As OLE_COLOR
Dim m_Font As Font
Dim m_Farbe_Urlaub As OLE_COLOR
Dim m_Schriftfarbe_urlaub As OLE_COLOR
Dim m_Farbe_HalberTag As OLE_COLOR
Dim m_Farbe_einzelner_Tag As OLE_COLOR
Dim m_Caption_Y_Achse As String
Dim m_TabFeld_ANZAHL_URLAUBSTAGE As String
Dim m_Quartal As Integer
Dim m_Jahr As Integer
'Dim m_Datenbankpfad As String
Dim m_Tabelle_DATUM As String
Dim m_Tabelle_NAMEN As String
Dim m_TabFeld_Datum_von As String
Dim m_TabFeld_Datum_bis As String
Dim m_TabFeld_Sonderurlaub As String
Dim m_TabFeld_NAMEN As String
Dim m_TabFeld_NUMMER As String
Dim m_Farbe_Sonderurlaub As OLE_COLOR
Dim m_BorderStyle As Integer
Dim m_Enabled As Boolean
Dim m_Nullpunkt_x As Long
'
'
''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
''MappingInfo=UserControl,UserControl,-1,BackColor
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = UserControl.BackColor
'End Property
'
'Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    UserControl.BackColor() = New_BackColor
'    PropertyChanged "BackColor"
'End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gibt einen Wert zurück, der bestimmt, ob ein Objekt auf vom Benutzer erzeugte Ereignisse reagieren kann, oder legt diesen fest."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property
Public Property Get Nullpunkt_x() As Long
    Nullpunkt_x = m_Nullpunkt_x
End Property

Public Property Let Nullpunkt_x(ByVal New_Nullpunkt_x As Long)
    m_Nullpunkt_x = New_Nullpunkt_x
    PropertyChanged "Nullpunkt_x"
End Property

Public Function show_chart() As Boolean
    Call MAKE_Y_ACHSE
    Select Case Quartal
        Case 1
            Call MAKE_X_ACHSE(0, 13)
        Case 2
            Call MAKE_X_ACHSE(13, 26)
        Case 3
            Call MAKE_X_ACHSE(26, 39)
        Case 4
            Call MAKE_X_ACHSE(39, 52)
        Case Else
            MsgBox "Falsche Angabe des Quartals. Wert muß zwischen 1 und 4 liegen."
            show_chart = False
            Exit Function
    End Select
    Call MAKE_GITTER
    Call PAINT_URLAUB
    LBL_UEBERSCHRIFT.Caption = "Angezeigter Zeitraum: " & Quartal & ". -tes Quartal " & Jahr & "."
    Frame_ALL.Width = Line_top.X2
    Frame_ALL.Height = LINE_S(0).Y2
    If LINE_S(0).Y2 > UserControl.Height Then
        VScroll.Left = UserControl.Width - VScroll.Width
        VScroll.Height = UserControl.Height
        VScroll.Min = 0
        VScroll.Max = 32767
        If (LINE_S(0).Y2 - Line_top.Y1) < 32767 Then VScroll.Max = LINE_S(0).Y2 - Line_top.Y1
        VScroll.Visible = True
    Else
        VScroll.Visible = False
    End If
    Label3.Caption = Left(Caption_Y_Achse, 18)
    'Setze die legende
    LBL_LEG_JAHRESURLAUB.BackColor = LBL_U(0).BackColor
    LBL_LEG_SONDERURLAUB.BackColor = Farbe_Sonderurlaub
    Shape_LEG_EIN_TAG.FillColor = Shape_EIN_TAG(0).FillColor
    Shape_LEG_HALBER_TAG.FillColor = Shape_HALBER_TAG(0).FillColor
    If Legende = True Then
        Frame_Legende.Visible = True
    Else
        Frame_Legende.Visible = False
    End If
    show_chart = True
End Function


Private Sub UserControl_Initialize()
    Frame_ALL.Width = Line_top.X2
    Frame_ALL.Height = LINE_S(0).Y2
    If LINE_S(0).Y2 > UserControl.Height Then
        VScroll.Left = UserControl.Width - VScroll.Width
        VScroll.Height = UserControl.Height
        VScroll.Min = 0
         VScroll.Max = 32767
        If (LINE_S(0).Y2 - Line_top.Y1) < 32767 Then VScroll.Max = LINE_S(0).Y2 - Line_top.Y1
        VScroll.Visible = True
        VScroll.Visible = True
    Else
        VScroll.Visible = False
    End If
End Sub

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    m_Nullpunkt_x = m_def_Nullpunkt_x
    m_BorderStyle = m_def_BorderStyle
    m_Farbe_Sonderurlaub = m_def_Farbe_Sonderurlaub
    m_Quartal = m_def_Quartal
    m_Jahr = m_def_Jahr
'    m_Datenbankpfad = m_def_Datenbankpfad
    m_Tabelle_DATUM = m_def_Tabelle_DATUM
    m_Tabelle_NAMEN = m_def_Tabelle_NAMEN
    m_TabFeld_Datum_von = m_def_TabFeld_Datum_von
    m_TabFeld_Datum_bis = m_def_TabFeld_Datum_bis
    m_TabFeld_Sonderurlaub = m_def_TabFeld_Sonderurlaub
    m_TabFeld_NAMEN = m_def_TabFeld_NAMEN
    m_TabFeld_NUMMER = m_def_TabFeld_NUMMER
    m_TabFeld_ANZAHL_URLAUBSTAGE = m_def_TabFeld_ANZAHL_URLAUBSTAGE
    m_Caption_Y_Achse = m_def_Caption_Y_Achse
'    m_ForeColor = m_def_ForeColor
    Set m_Font = Ambient.Font
    m_Farbe_Urlaub = m_def_Farbe_Urlaub
    m_Schriftfarbe_urlaub = m_def_Schriftfarbe_urlaub
    m_Farbe_HalberTag = m_def_Farbe_HalberTag
    m_Farbe_einzelner_Tag = m_def_Farbe_einzelner_Tag
    m_ForeColor = m_def_ForeColor
    m_Legende = m_def_Legende
    m_Filter_Feldname = m_def_Filter_Feldname
    m_ToolTip_Y = m_def_ToolTip_Y
    m_Filter_Bedingung = m_def_Filter_Bedingung
'    m_ConnectionString = m_def_ConnectionString
    m_ConnectionString = m_def_ConnectionString
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Line_KW(0).BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    m_Nullpunkt_x = PropBag.ReadProperty("Nullpunkt_x", m_def_Nullpunkt_x)
    Line_KW(0).BorderColor = PropBag.ReadProperty("ForeColor", -2147483640)
    LBL_U(0).BackColor = PropBag.ReadProperty("Farbe_Urlaub", &HC00000)
    LBL_U(0).ForeColor = PropBag.ReadProperty("Schriftfarbe_urlaub", &HC0FFFF)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Farbe_Sonderurlaub = PropBag.ReadProperty("Farbe_Sonderurlaub", m_def_Farbe_Sonderurlaub)
    m_Quartal = PropBag.ReadProperty("Quartal", m_def_Quartal)
    m_Jahr = PropBag.ReadProperty("Jahr", m_def_Jahr)
'    m_Datenbankpfad = PropBag.ReadProperty("Datenbankpfad", m_def_Datenbankpfad)
    m_Tabelle_DATUM = PropBag.ReadProperty("Tabelle_DATUM", m_def_Tabelle_DATUM)
    m_Tabelle_NAMEN = PropBag.ReadProperty("Tabelle_NAMEN", m_def_Tabelle_NAMEN)
    m_TabFeld_Datum_von = PropBag.ReadProperty("TabFeld_Datum_von", m_def_TabFeld_Datum_von)
    m_TabFeld_Datum_bis = PropBag.ReadProperty("TabFeld_Datum_bis", m_def_TabFeld_Datum_bis)
    m_TabFeld_Sonderurlaub = PropBag.ReadProperty("TabFeld_Sonderurlaub", m_def_TabFeld_Sonderurlaub)
    m_TabFeld_NAMEN = PropBag.ReadProperty("TabFeld_NAMEN", m_def_TabFeld_NAMEN)
    m_TabFeld_NUMMER = PropBag.ReadProperty("TabFeld_NUMMER", m_def_TabFeld_NUMMER)
    Shape_HALBER_TAG(0).FillColor = PropBag.ReadProperty("Farbe_HalberTag", &H8000&)
    Shape_EIN_TAG(0).FillColor = PropBag.ReadProperty("Farbe_einzelner_Tag", &HC00000)
    m_TabFeld_ANZAHL_URLAUBSTAGE = PropBag.ReadProperty("TabFeld_ANZAHL_URLAUBSTAGE", m_def_TabFeld_ANZAHL_URLAUBSTAGE)
    Frame_ALL.BackColor = PropBag.ReadProperty("Farbe_Hintergrund", &H8000000F)
    m_Caption_Y_Achse = PropBag.ReadProperty("Caption_Y_Achse", m_def_Caption_Y_Achse)
    Line_KW(0).BorderColor = PropBag.ReadProperty("ForeColor", -2147483640)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    LBL_U(0).BackColor = PropBag.ReadProperty("Farbe_Urlaub", &HC00000)
    LBL_U(0).ForeColor = PropBag.ReadProperty("Schriftfarbe_urlaub", &HC0FFFF)
    Shape_HALBER_TAG(0).FillColor = PropBag.ReadProperty("Farbe_HalberTag", &H8000&)
    Shape_EIN_TAG(0).FillColor = PropBag.ReadProperty("Farbe_einzelner_Tag", &HC00000)
'    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Farbe_Urlaub = PropBag.ReadProperty("Farbe_Urlaub", m_def_Farbe_Urlaub)
    m_Schriftfarbe_urlaub = PropBag.ReadProperty("Schriftfarbe_urlaub", m_def_Schriftfarbe_urlaub)
    m_Farbe_HalberTag = PropBag.ReadProperty("Farbe_HalberTag", m_def_Farbe_HalberTag)
    m_Farbe_einzelner_Tag = PropBag.ReadProperty("Farbe_einzelner_Tag", m_def_Farbe_einzelner_Tag)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Legende = PropBag.ReadProperty("Legende", m_def_Legende)
    m_Filter_Feldname = PropBag.ReadProperty("Filter_Feldname", m_def_Filter_Feldname)
    m_ToolTip_Y = PropBag.ReadProperty("ToolTip_Y", m_def_ToolTip_Y)
    m_Filter_Bedingung = PropBag.ReadProperty("Filter_Bedingung", m_def_Filter_Bedingung)
'    m_ConnectionString = PropBag.ReadProperty("ConnectionString", m_def_ConnectionString)
    m_ConnectionString = PropBag.ReadProperty("ConnectionString", m_def_ConnectionString)
End Sub

Private Sub UserControl_Resize()
    Frame_ALL.Width = Line_top.X2
    Frame_ALL.Height = LINE_S(0).Y2
    If LINE_S(0).Y2 > UserControl.Height Then
        VScroll.Left = UserControl.Width - VScroll.Width
        VScroll.Height = UserControl.Height
        VScroll.Min = 0
         VScroll.Max = 32767
        If (LINE_S(0).Y2 - Line_top.Y1) < 32767 Then VScroll.Max = LINE_S(0).Y2 - Line_top.Y1
        VScroll.Visible = True
        VScroll.Visible = True
    Else
        VScroll.Visible = False
    End If
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", Line_KW(0).BorderStyle, 1)
    Call PropBag.WriteProperty("Nullpunkt_x", m_Nullpunkt_x, m_def_Nullpunkt_x)
    Call PropBag.WriteProperty("ForeColor", Line_KW(0).BorderColor, -2147483640)
    Call PropBag.WriteProperty("Farbe_Urlaub", LBL_U(0).BackColor, &HC00000)
    Call PropBag.WriteProperty("Schriftfarbe_urlaub", LBL_U(0).ForeColor, &HC0FFFF)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Farbe_Sonderurlaub", m_Farbe_Sonderurlaub, m_def_Farbe_Sonderurlaub)
    Call PropBag.WriteProperty("Quartal", m_Quartal, m_def_Quartal)
    Call PropBag.WriteProperty("Jahr", m_Jahr, m_def_Jahr)
'    Call PropBag.WriteProperty("Datenbankpfad", m_Datenbankpfad, m_def_Datenbankpfad)
    Call PropBag.WriteProperty("Tabelle_DATUM", m_Tabelle_DATUM, m_def_Tabelle_DATUM)
    Call PropBag.WriteProperty("Tabelle_NAMEN", m_Tabelle_NAMEN, m_def_Tabelle_NAMEN)
    Call PropBag.WriteProperty("TabFeld_Datum_von", m_TabFeld_Datum_von, m_def_TabFeld_Datum_von)
    Call PropBag.WriteProperty("TabFeld_Datum_bis", m_TabFeld_Datum_bis, m_def_TabFeld_Datum_bis)
    Call PropBag.WriteProperty("TabFeld_Sonderurlaub", m_TabFeld_Sonderurlaub, m_def_TabFeld_Sonderurlaub)
    Call PropBag.WriteProperty("TabFeld_NAMEN", m_TabFeld_NAMEN, m_def_TabFeld_NAMEN)
    Call PropBag.WriteProperty("TabFeld_NUMMER", m_TabFeld_NUMMER, m_def_TabFeld_NUMMER)
    Call PropBag.WriteProperty("Farbe_HalberTag", Shape_HALBER_TAG(0).FillColor, &H8000&)
    Call PropBag.WriteProperty("Farbe_einzelner_Tag", Shape_EIN_TAG(0).FillColor, &HC00000)
    Call PropBag.WriteProperty("TabFeld_ANZAHL_URLAUBSTAGE", m_TabFeld_ANZAHL_URLAUBSTAGE, m_def_TabFeld_ANZAHL_URLAUBSTAGE)
    Call PropBag.WriteProperty("Farbe_Hintergrund", Frame_ALL.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Caption_Y_Achse", m_Caption_Y_Achse, m_def_Caption_Y_Achse)
    Call PropBag.WriteProperty("ForeColor", Line_KW(0).BorderColor, -2147483640)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("Farbe_Urlaub", LBL_U(0).BackColor, &HC00000)
    Call PropBag.WriteProperty("Schriftfarbe_urlaub", LBL_U(0).ForeColor, &HC0FFFF)
    Call PropBag.WriteProperty("Farbe_HalberTag", Shape_HALBER_TAG(0).FillColor, &H8000&)
    Call PropBag.WriteProperty("Farbe_einzelner_Tag", Shape_EIN_TAG(0).FillColor, &HC00000)
'    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("Farbe_Urlaub", m_Farbe_Urlaub, m_def_Farbe_Urlaub)
    Call PropBag.WriteProperty("Schriftfarbe_urlaub", m_Schriftfarbe_urlaub, m_def_Schriftfarbe_urlaub)
    Call PropBag.WriteProperty("Farbe_HalberTag", m_Farbe_HalberTag, m_def_Farbe_HalberTag)
    Call PropBag.WriteProperty("Farbe_einzelner_Tag", m_Farbe_einzelner_Tag, m_def_Farbe_einzelner_Tag)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Legende", m_Legende, m_def_Legende)
    Call PropBag.WriteProperty("Filter_Feldname", m_Filter_Feldname, m_def_Filter_Feldname)
    Call PropBag.WriteProperty("ToolTip_Y", m_ToolTip_Y, m_def_ToolTip_Y)
    Call PropBag.WriteProperty("Filter_Bedingung", m_Filter_Bedingung, m_def_Filter_Bedingung)
'    Call PropBag.WriteProperty("ConnectionString", m_ConnectionString, m_def_ConnectionString)
    Call PropBag.WriteProperty("ConnectionString", m_ConnectionString, m_def_ConnectionString)
End Sub
Public Sub MAKE_Y_ACHSE()
    Dim i As Integer
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    Dim SQL_TEXT As String
    SQL_TEXT = ""
    i = 0
    On Error Resume Next
    Load LINE_W(1)
    If Err.Number = 360 Then
        For i = 1 To ANZAHL_MA + 1
            Unload LINE_W(i)
            Unload LBL_U(i)
            Unload LBL_Y(i)
            Unload LBL_NR(i)
        Next i
    Else
        Unload LINE_W(1)
    End If
    On Error GoTo 0
    On Error GoTo Fehler
    'Beispiel für die Eigenschaft ConnectionString:
    '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sample.mdb;Persist Security Info=False"
    conn.Open ConnectionString
    SQL_TEXT = "Select Distinct " & TabFeld_NAMEN & ", " & TabFeld_NUMMER & ""
        If TabFeld_NAMEN = TabFeld_NUMMER Then SQL_TEXT = "Select Distinct " & TabFeld_NAMEN
        If Trim(ToolTip_Y) <> "" Then SQL_TEXT = SQL_TEXT + ", " + ToolTip_Y
    If Trim(Filter_Feldname) <> "" Then
        'Welcher Datentyp ist das Feld Filter_Feldname ?
        rs.Open "Select " & Filter_Feldname & " from " & Tabelle_NAMEN & "", conn, adOpenKeyset, adLockOptimistic, adCmdText
        Select Case rs.Fields(Filter_Feldname).Type
            Case dbText
            rs.Close
            rs.Open SQL_TEXT & " FROM " & Tabelle_NAMEN & " Where (" & Filter_Feldname & " = '" & Filter_Bedingung & "') order by " & TabFeld_NAMEN & "", conn, adOpenKeyset, adLockOptimistic, adCmdText
            Case dbLong
            rs.Close
            rs.Open SQL_TEXT & " FROM " & Tabelle_NAMEN & " Where (" & Filter_Feldname & " = " & Filter_Bedingung & ") order by " & TabFeld_NAMEN & "", conn, adOpenKeyset, adLockOptimistic, adCmdText
            Case Else
            rs.Close
            rs.Open SQL_TEXT & " FROM " & Tabelle_NAMEN & " Where (" & Filter_Feldname & " = " & Filter_Bedingung & ") order by " & TabFeld_NAMEN & "", conn, adOpenKeyset, adLockOptimistic, adCmdText
        End Select
    Else
        rs.Open SQL_TEXT & " FROM " & Tabelle_NAMEN & " order by " & TabFeld_NAMEN & "", conn, adOpenKeyset, adLockOptimistic, adCmdText
    End If
    i = 0
    While Not rs.EOF
        i = i + 1
        Load LINE_W(i)
        LINE_W(i).X1 = LINE_W(i - 1).X1
        LINE_W(i).X2 = LINE_W(i - 1).X2
        LINE_W(i).Y1 = LINE_W(i - 1).Y1 + 350
        LINE_W(i).Y2 = LINE_W(i - 1).Y2 + 350
        LINE_W(i).Visible = True
        
        Load LBL_Y(i)
        LBL_Y(i).Left = LBL_Y(i - 1).Left
        LBL_Y(i).Top = LINE_W(i).Y2 - 300
        LAENGE_Y_ACHSE = LBL_Y(i).Top
        LBL_Y(i).Caption = UCase(Left(rs(TabFeld_NAMEN), 18))
        LBL_Y(i).Visible = True
        If Trim(ToolTip_Y) <> "" Then LBL_Y(i).ToolTipText = rs(ToolTip_Y)
        'versteckte Labels für MA-NR
        Load LBL_NR(i)
        LBL_NR(i).Caption = rs(TabFeld_NUMMER)
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    LINE_S(0).Y2 = (LAENGE_Y_ACHSE + 350) - (LINE_W(i).Y2 - LBL_Y(i).Top - LBL_Y(i).Height)
    ANZAHL_MA = i
    Exit Sub
Fehler:
    If Err.Number = -2147467259 Then
        MsgBox "Die Datenbank " & DB_PFAD & " wurde nicht gefunden." & Chr(13) & Chr(13) & "" _
        & "Stellen Sie sicher, dass Sie im angebenen Pfad existiert oder installieren Sie ContactPro neu.", vbOKOnly + vbExclamation, "Datenbank ??"
        Exit Sub
    Else
        MsgBox "Fehler: " & Err.Number & Chr(13) & "Beschreibung: " & Chr(13) & Err.Description, vbOKOnly + vbExclamation
        Exit Sub
    End If
End Sub
Public Sub MAKE_X_ACHSE(ByVal VON, BIS As Integer)
    Dim i As Integer
    Dim KW_SPACER As Long
    KW_SPACER = ((Width - 100) - LBL_X(0).Left) / 14
    'KW_SPACER/7 = 1 Tag
    On Error Resume Next
    Load LBL_X(1)
    If Err.Number = 360 Then
        For i = 1 To 14
            Unload LBL_X(i)
        Next i
    Else
        Unload LBL_X(1)
    End If
    On Error GoTo Fehler
    For i = 0 To 13
        Load LBL_X(i + 1)
        If (i + 1) = 1 Then
            LBL_X(i + 1).Top = LBL_X(i).Top
            LBL_X(i + 1).Left = LBL_X(i).Left
            LBL_X(i + 1).Caption = VON + i
            If (VON + i) < 10 Then LBL_X(i + 1).Caption = "0" + CStr(VON + i)
        Else
            LBL_X(i + 1).Top = LBL_X(i).Top
            LBL_X(i + 1).Left = LBL_X(i).Left + KW_SPACER
            LBL_X(i + 1).Caption = VON + i
            If (VON + i) < 10 Then LBL_X(i + 1).Caption = "0" + CStr(VON + i)
        End If
        LBL_X(i + 1).AutoSize = True
        LBL_X(i + 1).Visible = False
        LAENGE_X_ACHSE = LBL_X(i + 1).Left + (LBL_X(i + 1).Width / 2)
    Next i
    
    For i = 0 To ANZAHL_MA
        LINE_W(i).X2 = LAENGE_X_ACHSE
    Next i
    Line_top.X2 = LAENGE_X_ACHSE
    Exit Sub
Fehler:
    MsgBox "Fehler: " & Err.Number & Chr(13) & "Beschreibung: " & Chr(13) & Err.Description, vbOKOnly + vbExclamation
    Exit Sub
End Sub
Public Sub PAINT_URLAUB()
    'ermittel hier für einen Mitarbeiter die Urlaubsspanne
    'gebe die Koordinaten zur Linienzeichnung zurück
    Dim i As Integer
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim KWOCHE As Integer
    Dim THIS_MA As Variant
    Dim TAGE_IN_TWIPS As Long
    Dim FIRSTDAY As String
    Dim U_VON_IN_TAGEN As Integer
    Dim U_BIS_IN_TAGEN As Integer
    Dim POS_U, POS_H, POS_E As Integer
    Dim SQL_1 As String
    Dim CONTROL_MAX As Integer
    CONTROL_MAX = 100
    SQL_1 = ""
    POS_U = 0
    POS_H = 0
    POS_E = 0
    FIRSTDAY = "01.01." + CStr(Jahr)
    'ist das ein Montag ?
    While Not Weekday(CDate(FIRSTDAY), vbMonday) = 1
        FIRSTDAY = CStr(DateAdd("d", 1, CDate(FIRSTDAY)))
    Wend
    i = 0
    On Error Resume Next
    'Load LBL_U(1)
    'If Err.Number = 360 Then
        For i = 1 To CONTROL_MAX
        'For i = 1 To ANZAHL_LINIEN_URLAUB + 1
            Unload LBL_U(i)
            Unload Shape_EIN_TAG(i)
            Unload Shape_HALBER_TAG(i)
        Next i
    'Else
        Unload LBL_U(1)
    'End If
    Load Shape_HALBER_TAG(1)
    If Err.Number = 360 Then
        For i = 1 To ANZAHL_HALBE_TAGE + 1
            Unload Shape_HALBER_TAG(i)
        Next i
    Else
        Unload Shape_HALBER_TAG(1)
    End If
    Load Shape_EIN_TAG(1)
    If Err.Number = 360 Then
        For i = 1 To ANZAHL_EINZELNE_TAGE + 1
            Unload Shape_EIN_TAG(i)
        Next i
    Else
        Unload Shape_EIN_TAG(1)
    End If
    On Error GoTo 0
    'On Error GoTo Fehler
    'wieviele TWIPS entsprechen den aktuell gezeigten 13 KW ?
    TAGE_IN_TWIPS = (Line_KW(14).X1 - LINE_S(0).X1) / (13 * 7)
    conn.Open ConnectionString
    SQL_1 = TabFeld_NUMMER & ", " & TabFeld_Datum_von & ", " & TabFeld_Datum_bis
    If TabFeld_Sonderurlaub <> "" Then SQL_1 = TabFeld_Sonderurlaub & ", " & SQL_1
    If TabFeld_ANZAHL_URLAUBSTAGE <> "" Then SQL_1 = TabFeld_ANZAHL_URLAUBSTAGE & ", " & SQL_1
    rs.Open "Select " & SQL_1 & "" _
        & " FROM " & Tabelle_DATUM & " WHERE (year(" & TabFeld_Datum_von & ")= " & Jahr & ") order by " & TabFeld_NUMMER & "", conn, adOpenKeyset, adLockOptimistic, adCmdText
    'Für alle angezeigten Mitarbeiter
    For i = 1 To ANZAHL_MA
        rs.MoveFirst
        THIS_MA = LBL_NR(i).Caption
        'wenn vom Typ Long
        If rs.Fields(TabFeld_NUMMER).Type = adInteger Then rs.Find "" & TabFeld_NUMMER & "=" & THIS_MA, 0, adSearchForward
        'wenn vom Typ String
        If rs.Fields(TabFeld_NUMMER).Type = adChar Then rs.Find "" & TabFeld_NUMMER & "='" & THIS_MA & "'", 0, adSearchForward
        If Not rs.EOF Then
            KWOCHE = DatePart("ww", rs(TabFeld_Datum_von), vbMonday, vbFirstFullWeek)
            If CHECK_QUARTAL(KWOCHE) = True Then
                U_VON_IN_TAGEN = DateDiff("d", CDate(FIRSTDAY), rs(TabFeld_Datum_von), vbMonday, vbFirstFullWeek)
                U_BIS_IN_TAGEN = DateDiff("d", CDate(FIRSTDAY), rs(TabFeld_Datum_bis), vbMonday, vbFirstFullWeek)
                U_BIS_IN_TAGEN = U_BIS_IN_TAGEN - U_VON_IN_TAGEN
                'Ein ganzer Tag, ein Halber oder mehr als einer ?
                Select Case U_BIS_IN_TAGEN
                    Case Is > 0
                        'Positioniere Starttag abhängig vom Quartal
                        U_VON_IN_TAGEN = U_VON_IN_TAGEN - (13 * 7 * (Quartal - 1))
                        On Error Resume Next
                        Load LBL_U(POS_U + 1)
                        On Error GoTo 0
                        LBL_U(POS_U + 1).Left = (U_VON_IN_TAGEN * TAGE_IN_TWIPS) + LINE_S(0).X1
                        LBL_U(POS_U + 1).Width = ((U_BIS_IN_TAGEN + 1) * TAGE_IN_TWIPS)
                        LBL_U(POS_U + 1).Top = (LBL_Y(i).Top)
                        LBL_U(POS_U + 1).Caption = U_BIS_IN_TAGEN + 1
                        LBL_U(POS_U + 1).BackColor = Farbe_Urlaub
                        LBL_U(POS_U + 1).ForeColor = Schriftfarbe_urlaub
                        If TabFeld_Sonderurlaub <> "" Then
                        If rs(TabFeld_Sonderurlaub) = True Then LBL_U(POS_U + 1).BackColor = Farbe_Sonderurlaub
                        End If
                        LBL_U(POS_U + 1).Visible = True
                        LBL_U(POS_U + 1).ToolTipText = CStr(rs(TabFeld_Datum_von)) + " bis " + CStr(rs(TabFeld_Datum_bis))
                        POS_U = POS_U + 1
                    Case Is = 0
                        If TabFeld_ANZAHL_URLAUBSTAGE <> "" Then
                            If rs(TabFeld_ANZAHL_URLAUBSTAGE) < 1 Then
                                U_VON_IN_TAGEN = U_VON_IN_TAGEN - (13 * 7 * (Quartal - 1))
                                Load Shape_HALBER_TAG(POS_H + 1)
                                Shape_HALBER_TAG(POS_H + 1).Left = (U_VON_IN_TAGEN * TAGE_IN_TWIPS) + LINE_S(0).X1
                                Shape_HALBER_TAG(POS_H + 1).Top = (LBL_Y(i).Top)
                                Shape_HALBER_TAG(POS_H + 1).FillColor = Farbe_HalberTag
                                Shape_HALBER_TAG(POS_H + 1).Visible = True
                                POS_H = POS_H + 1
                            Else
                                U_VON_IN_TAGEN = U_VON_IN_TAGEN - (13 * 7 * (Quartal - 1))
                                Load Shape_EIN_TAG(POS_E + 1)
                                Shape_EIN_TAG(POS_E + 1).Left = (U_VON_IN_TAGEN * TAGE_IN_TWIPS) + LINE_S(0).X1
                                Shape_EIN_TAG(POS_E + 1).Top = (LBL_Y(i).Top)
                                Shape_EIN_TAG(POS_E + 1).FillColor = Farbe_einzelner_Tag
                                Shape_EIN_TAG(POS_E + 1).Visible = True
                                POS_E = POS_E + 1
                            End If
                        Else
                            U_VON_IN_TAGEN = U_VON_IN_TAGEN - (13 * 7 * (Quartal - 1))
                            Load Shape_EIN_TAG(POS_E + 1)
                            Shape_EIN_TAG(POS_E + 1).Left = (U_VON_IN_TAGEN * TAGE_IN_TWIPS) + LINE_S(0).X1
                            Shape_EIN_TAG(POS_E + 1).Top = (LBL_Y(i).Top)
                            Shape_EIN_TAG(POS_E + 1).FillColor = Farbe_einzelner_Tag
                            Shape_EIN_TAG(POS_E + 1).Visible = True
                            POS_E = POS_E + 1
                        End If
                    End Select
            End If
        End If
        While rs.EOF = False
            'wenn vom Typ Long
            If rs.Fields(TabFeld_NUMMER).Type = adInteger Then rs.Find "" & TabFeld_NUMMER & "=" & THIS_MA, 0, adSearchForward
            'wenn vom Typ String
            If rs.Fields(TabFeld_NUMMER).Type = adChar Then rs.Find "" & TabFeld_NUMMER & "='" & THIS_MA & "'", 0, adSearchForward
            If Not rs.EOF Then
                KWOCHE = DatePart("ww", rs(TabFeld_Datum_von), vbMonday, vbFirstFullWeek)
                If CHECK_QUARTAL(KWOCHE) = True Then
                    U_VON_IN_TAGEN = DateDiff("d", CDate(FIRSTDAY), rs(TabFeld_Datum_von), vbMonday, vbFirstFullWeek)
                    U_BIS_IN_TAGEN = DateDiff("d", CDate(FIRSTDAY), rs(TabFeld_Datum_bis), vbMonday, vbFirstFullWeek)
                    U_BIS_IN_TAGEN = U_BIS_IN_TAGEN - U_VON_IN_TAGEN
                    'Ein ganzer Tag, ein Halber oder mehr als einer ?
                Select Case U_BIS_IN_TAGEN
                    Case Is > 0
                        'Positioniere Starttag abhängig vom Quartal
                        U_VON_IN_TAGEN = U_VON_IN_TAGEN - (13 * 7 * (Quartal - 1))
                        On Error Resume Next
                        Load LBL_U(POS_U + 1)
                        On Error GoTo 0
                        LBL_U(POS_U + 1).Left = (U_VON_IN_TAGEN * TAGE_IN_TWIPS) + LINE_S(0).X1
                        LBL_U(POS_U + 1).Width = ((U_BIS_IN_TAGEN + 1) * TAGE_IN_TWIPS)
                        LBL_U(POS_U + 1).Top = (LBL_Y(i).Top)
                        LBL_U(POS_U + 1).Caption = U_BIS_IN_TAGEN + 1
                        LBL_U(POS_U + 1).BackColor = Farbe_Urlaub
                        LBL_U(POS_U + 1).ForeColor = Schriftfarbe_urlaub
                        If TabFeld_Sonderurlaub <> "" Then
                        If rs(TabFeld_Sonderurlaub) = True Then LBL_U(POS_U + 1).BackColor = Farbe_Sonderurlaub
                        End If
                        LBL_U(POS_U + 1).Visible = True
                        LBL_U(POS_U + 1).ToolTipText = CStr(rs(TabFeld_Datum_von)) + " bis " + CStr(rs(TabFeld_Datum_bis))
                        POS_U = POS_U + 1
                    Case Is = 0
                        If TabFeld_ANZAHL_URLAUBSTAGE <> "" Then
                            If rs(TabFeld_ANZAHL_URLAUBSTAGE) < 1 Then
                                U_VON_IN_TAGEN = U_VON_IN_TAGEN - (13 * 7 * (Quartal - 1))
                                Load Shape_HALBER_TAG(POS_H + 1)
                                Shape_HALBER_TAG(POS_H + 1).Left = (U_VON_IN_TAGEN * TAGE_IN_TWIPS) + LINE_S(0).X1
                                Shape_HALBER_TAG(POS_H + 1).Top = (LBL_Y(i).Top)
                                Shape_HALBER_TAG(POS_H + 1).FillColor = Farbe_HalberTag
                                Shape_HALBER_TAG(POS_H + 1).Visible = True
                                POS_H = POS_H + 1
                            Else
                                U_VON_IN_TAGEN = U_VON_IN_TAGEN - (13 * 7 * (Quartal - 1))
                                Load Shape_EIN_TAG(POS_E + 1)
                                Shape_EIN_TAG(POS_E + 1).Left = (U_VON_IN_TAGEN * TAGE_IN_TWIPS) + LINE_S(0).X1
                                Shape_EIN_TAG(POS_E + 1).Top = (LBL_Y(i).Top)
                                Shape_EIN_TAG(POS_E + 1).FillColor = Farbe_einzelner_Tag
                                Shape_EIN_TAG(POS_E + 1).Visible = True
                                POS_E = POS_E + 1
                            End If
                        Else
                            U_VON_IN_TAGEN = U_VON_IN_TAGEN - (13 * 7 * (Quartal - 1))
                            Load Shape_EIN_TAG(POS_E + 1)
                            Shape_EIN_TAG(POS_E + 1).Left = (U_VON_IN_TAGEN * TAGE_IN_TWIPS) + LINE_S(0).X1
                            Shape_EIN_TAG(POS_E + 1).Top = (LBL_Y(i).Top)
                            Shape_EIN_TAG(POS_E + 1).FillColor = Farbe_einzelner_Tag
                            Shape_EIN_TAG(POS_E + 1).Visible = True
                            POS_E = POS_E + 1
                        End If
                    End Select
                End If
            End If
            On Error Resume Next
            rs.MoveNext
            On Error GoTo 0
        Wend
    Next i
    ANZAHL_LINIEN_URLAUB = POS_U + 1
    ANZAHL_HALBE_TAGE = POS_H + 1
    ANZAHL_EINZELNE_TAGE = POS_E + 1
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    Exit Sub
Fehler:
    If Err.Number = -2147467259 Then
        MsgBox "Die Datenbank " & DB_PFAD & " wurde nicht gefunden." & Chr(13) & Chr(13) & "" _
        & "Stellen Sie sicher, dass Sie im angebenen Pfad existiert oder installieren Sie ContactPro neu.", vbOKOnly + vbExclamation, "Datenbank ??"
        Exit Sub
    Else
        MsgBox "Fehler: " & Err.Number & Chr(13) & "Beschreibung: " & Chr(13) & Err.Description, vbOKOnly + vbExclamation
        Exit Sub
    End If
End Sub
Public Function CHECK_QUARTAL(KWOCHE As Integer) As Boolean
    'Urlaubslinien nur zeigen, wenn Urlaub auch im Quartal der angeklickten
    'Schaltfläche beginnt
    Dim U_QUARTAL_START As Integer
    On Error GoTo Fehler
    If KWOCHE >= 1 And KWOCHE <= 13 Then U_QUARTAL_START = 1
    If KWOCHE >= 13 And KWOCHE <= 26 Then U_QUARTAL_START = 2
    If KWOCHE >= 26 And KWOCHE <= 39 Then U_QUARTAL_START = 3
    If KWOCHE >= 39 And KWOCHE <= 52 Then U_QUARTAL_START = 4
    CHECK_QUARTAL = False
    If U_QUARTAL_START = Quartal Then CHECK_QUARTAL = True
    Exit Function
Fehler:
    MsgBox "Fehler: " & Err.Number & Chr(13) & "Beschreibung: " & Chr(13) & Err.Description, vbOKOnly + vbExclamation
    Exit Function
End Function
Public Sub MAKE_GITTER()
Dim i As Integer
    On Error Resume Next
    For i = 1 To 14
        Unload Line_KW(i)
        Unload LBL_HEADER(i)
    Next i
    On Error GoTo 0
    On Error GoTo Fehler
    For i = 1 To 14
        Load Line_KW(i)
        Line_KW(i).X1 = LBL_X(i).Left + (LBL_X(i).Width / 2)
        Line_KW(i).X2 = LBL_X(i).Left + (LBL_X(i).Width / 2)
        Line_KW(i).Y1 = Line_top.Y1
        Line_KW(i).Y2 = LINE_S(0).Y2
        Line_KW(i).Visible = True
    Next i
    Line_KW(1).Visible = False
    'Nun die Spaltenheader der KW
    Load LBL_HEADER(1)
    LBL_HEADER(1).Left = LINE_S(0).X2
    LBL_HEADER(1).Top = LBL_X(0).Top
    LBL_HEADER(1).Width = ((Width - 100) - LBL_X(0).Left) / 14
    LBL_HEADER(1).Caption = LBL_X(2).Caption
    LBL_HEADER(1).Visible = True
    For i = 2 To 14
        Load LBL_HEADER(i)
        LBL_HEADER(i).Left = Line_KW(i - 1).X2
        LBL_HEADER(i).Top = LBL_X(0).Top
        LBL_HEADER(i).Width = ((Width - 100) - LBL_X(0).Left) / 14
        LBL_HEADER(i).Caption = LBL_X(i).Caption
        LBL_HEADER(i).Visible = True
    Next i
    Exit Sub
Fehler:
    MsgBox "Fehler: " & Err.Number & Chr(13) & "Beschreibung: " & Chr(13) & Err.Description, vbOKOnly + vbExclamation
    Exit Sub
End Sub


Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Gibt den Rahmenstil für ein Objekt zurück oder legt diesen fest."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property



Public Property Get Farbe_Sonderurlaub() As OLE_COLOR
    Farbe_Sonderurlaub = m_Farbe_Sonderurlaub
End Property

Public Property Let Farbe_Sonderurlaub(ByVal New_Farbe_Sonderurlaub As OLE_COLOR)
    m_Farbe_Sonderurlaub = New_Farbe_Sonderurlaub
    PropertyChanged "Farbe_Sonderurlaub"
End Property

Public Property Get Quartal() As Integer
Attribute Quartal.VB_ProcData.VB_Invoke_Property = "Zeitraum"
    Quartal = m_Quartal
End Property

Public Property Let Quartal(ByVal New_Quartal As Integer)
    m_Quartal = New_Quartal
    PropertyChanged "Quartal"
End Property

Public Property Get Jahr() As Integer
Attribute Jahr.VB_ProcData.VB_Invoke_Property = "Zeitraum"
    Jahr = m_Jahr
End Property

Public Property Let Jahr(ByVal New_Jahr As Integer)
    m_Jahr = New_Jahr
    PropertyChanged "Jahr"
End Property
'
'Public Property Get Datenbankpfad() As String
'    Datenbankpfad = m_Datenbankpfad
'End Property
'
'Public Property Let Datenbankpfad(ByVal New_Datenbankpfad As String)
'    m_Datenbankpfad = New_Datenbankpfad
'    PropertyChanged "Datenbankpfad"
'End Property

Public Property Get Tabelle_DATUM() As String
Attribute Tabelle_DATUM.VB_ProcData.VB_Invoke_Property = "Datenbank"
    Tabelle_DATUM = m_Tabelle_DATUM
End Property

Public Property Let Tabelle_DATUM(ByVal New_Tabelle_DATUM As String)
    m_Tabelle_DATUM = New_Tabelle_DATUM
    PropertyChanged "Tabelle_DATUM"
End Property

Public Property Get Tabelle_NAMEN() As String
Attribute Tabelle_NAMEN.VB_ProcData.VB_Invoke_Property = "Datenbank"
    Tabelle_NAMEN = m_Tabelle_NAMEN
End Property

Public Property Let Tabelle_NAMEN(ByVal New_Tabelle_NAMEN As String)
    m_Tabelle_NAMEN = New_Tabelle_NAMEN
    PropertyChanged "Tabelle_NAMEN"
End Property

Public Property Get TabFeld_Datum_von() As String
Attribute TabFeld_Datum_von.VB_ProcData.VB_Invoke_Property = "Datenbank"
    TabFeld_Datum_von = m_TabFeld_Datum_von
End Property

Public Property Let TabFeld_Datum_von(ByVal New_TabFeld_Datum_von As String)
    m_TabFeld_Datum_von = New_TabFeld_Datum_von
    PropertyChanged "TabFeld_Datum_von"
End Property

Public Property Get TabFeld_Datum_bis() As String
Attribute TabFeld_Datum_bis.VB_ProcData.VB_Invoke_Property = "Datenbank"
    TabFeld_Datum_bis = m_TabFeld_Datum_bis
End Property

Public Property Let TabFeld_Datum_bis(ByVal New_TabFeld_Datum_bis As String)
    m_TabFeld_Datum_bis = New_TabFeld_Datum_bis
    PropertyChanged "TabFeld_Datum_bis"
End Property

Public Property Get TabFeld_Sonderurlaub() As String
Attribute TabFeld_Sonderurlaub.VB_ProcData.VB_Invoke_Property = "Datenbank"
    TabFeld_Sonderurlaub = m_TabFeld_Sonderurlaub
End Property

Public Property Let TabFeld_Sonderurlaub(ByVal New_TabFeld_Sonderurlaub As String)
    m_TabFeld_Sonderurlaub = New_TabFeld_Sonderurlaub
    PropertyChanged "TabFeld_Sonderurlaub"
End Property

Public Property Get TabFeld_NAMEN() As String
Attribute TabFeld_NAMEN.VB_ProcData.VB_Invoke_Property = "Datenbank"
    TabFeld_NAMEN = m_TabFeld_NAMEN
End Property

Public Property Let TabFeld_NAMEN(ByVal New_TabFeld_NAMEN As String)
    m_TabFeld_NAMEN = New_TabFeld_NAMEN
    PropertyChanged "TabFeld_NAMEN"
End Property

Public Property Get TabFeld_NUMMER() As String
Attribute TabFeld_NUMMER.VB_ProcData.VB_Invoke_Property = "Datenbank"
    TabFeld_NUMMER = m_TabFeld_NUMMER
End Property

Public Property Let TabFeld_NUMMER(ByVal New_TabFeld_NUMMER As String)
    m_TabFeld_NUMMER = New_TabFeld_NUMMER
    PropertyChanged "TabFeld_NUMMER"
End Property


Public Property Get TabFeld_ANZAHL_URLAUBSTAGE() As String
Attribute TabFeld_ANZAHL_URLAUBSTAGE.VB_ProcData.VB_Invoke_Property = "Datenbank"
    TabFeld_ANZAHL_URLAUBSTAGE = m_TabFeld_ANZAHL_URLAUBSTAGE
End Property

Public Property Let TabFeld_ANZAHL_URLAUBSTAGE(ByVal New_TabFeld_ANZAHL_URLAUBSTAGE As String)
    m_TabFeld_ANZAHL_URLAUBSTAGE = New_TabFeld_ANZAHL_URLAUBSTAGE
    PropertyChanged "TabFeld_ANZAHL_URLAUBSTAGE"
End Property

Private Sub VScroll_Change()
    Frame_ALL.Top = -VScroll.Value
End Sub

Private Sub VScroll_Scroll()
    Frame_ALL.Top = -VScroll.Value
End Sub
'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Erzwingt ein vollständiges Neuzeichnen eines Objekts."
    UserControl.Refresh
End Sub
'
''
''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
''MappingInfo=UserControl,UserControl,-1,BackColor
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = UserControl.BackColor
'End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=Frame_ALL,Frame_ALL,-1,BackColor
Public Property Get Farbe_Hintergrund() As OLE_COLOR
Attribute Farbe_Hintergrund.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."
    Farbe_Hintergrund = Frame_ALL.BackColor
End Property

Public Property Let Farbe_Hintergrund(ByVal New_Farbe_Hintergrund As OLE_COLOR)
    Frame_ALL.BackColor() = New_Farbe_Hintergrund
    PropertyChanged "Farbe_Hintergrund"
End Property

Public Property Get Caption_Y_Achse() As String
Attribute Caption_Y_Achse.VB_ProcData.VB_Invoke_Property = "Zeitraum"
    Caption_Y_Achse = m_Caption_Y_Achse
End Property

Public Property Let Caption_Y_Achse(ByVal New_Caption_Y_Achse As String)
    m_Caption_Y_Achse = New_Caption_Y_Achse
    PropertyChanged "Caption_Y_Achse"
End Property
''
'''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'''MappingInfo=Line_KW(0),Line_KW,0,BorderColor
''Public Property Get ForeColor() As Long
''    ForeColor = Line_KW(0).BorderColor
''End Property
''
''Public Property Let ForeColor(ByVal New_ForeColor As Long)
''    Line_KW(0).BorderColor() = New_ForeColor
''    PropertyChanged "ForeColor"
''End Property
''
'''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'''MappingInfo=LBL_U(0),LBL_U,0,Font
''Public Property Get Font() As Font
''    Set Font = LBL_U(0).Font
''End Property
''
''Public Property Set Font(ByVal New_Font As Font)
''    Set LBL_U(0).Font = New_Font
''    PropertyChanged "Font"
''End Property
''
'''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'''MappingInfo=LBL_U(0),LBL_U,0,BackColor
''Public Property Get Farbe_Urlaub() As OLE_COLOR
''    Farbe_Urlaub = LBL_U(0).BackColor
''End Property
''
''Public Property Let Farbe_Urlaub(ByVal New_Farbe_Urlaub As OLE_COLOR)
''    LBL_U(0).BackColor() = New_Farbe_Urlaub
''    PropertyChanged "Farbe_Urlaub"
''End Property
''
'''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'''MappingInfo=LBL_U(0),LBL_U,0,ForeColor
''Public Property Get Schriftfarbe_urlaub() As OLE_COLOR
''    Schriftfarbe_urlaub = LBL_U(0).ForeColor
''End Property
''
''Public Property Let Schriftfarbe_urlaub(ByVal New_Schriftfarbe_urlaub As OLE_COLOR)
''    LBL_U(0).ForeColor() = New_Schriftfarbe_urlaub
''    PropertyChanged "Schriftfarbe_urlaub"
''End Property
''
'''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'''MappingInfo=Shape_HALBER_TAG(0),Shape_HALBER_TAG,0,FillColor
''Public Property Get Farbe_HalberTag() As OLE_COLOR
''    Farbe_HalberTag = Shape_HALBER_TAG(0).FillColor
''End Property
''
''Public Property Let Farbe_HalberTag(ByVal New_Farbe_HalberTag As OLE_COLOR)
''    Shape_HALBER_TAG(0).FillColor() = New_Farbe_HalberTag
''    PropertyChanged "Farbe_HalberTag"
''End Property
''
'''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'''MappingInfo=Shape_EIN_TAG(0),Shape_EIN_TAG,0,FillColor
''Public Property Get Farbe_einzelner_Tag() As OLE_COLOR
''    Farbe_einzelner_Tag = Shape_EIN_TAG(0).FillColor
''End Property
''
''Public Property Let Farbe_einzelner_Tag(ByVal New_Farbe_einzelner_Tag As OLE_COLOR)
''    Shape_EIN_TAG(0).FillColor() = New_Farbe_einzelner_Tag
''    PropertyChanged "Farbe_einzelner_Tag"
''End Property
''
'Public Property Get ForeColor() As OLE_COLOR
'    ForeColor = m_ForeColor
'End Property
'
'Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
'    m_ForeColor = New_ForeColor
'    PropertyChanged "ForeColor"
'End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Gibt ein Font-Objekt zurück."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get Farbe_Urlaub() As OLE_COLOR
Attribute Farbe_Urlaub.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."
    Farbe_Urlaub = m_Farbe_Urlaub
End Property

Public Property Let Farbe_Urlaub(ByVal New_Farbe_Urlaub As OLE_COLOR)
    m_Farbe_Urlaub = New_Farbe_Urlaub
    PropertyChanged "Farbe_Urlaub"
End Property

Public Property Get Schriftfarbe_urlaub() As OLE_COLOR
Attribute Schriftfarbe_urlaub.VB_Description = "Gibt die Vordergrundfarbe zurück, die zum Anzeigen von Text und Grafiken in einem Objekt verwendet wird, oder legt diese fest."
    Schriftfarbe_urlaub = m_Schriftfarbe_urlaub
End Property

Public Property Let Schriftfarbe_urlaub(ByVal New_Schriftfarbe_urlaub As OLE_COLOR)
    m_Schriftfarbe_urlaub = New_Schriftfarbe_urlaub
    PropertyChanged "Schriftfarbe_urlaub"
End Property

Public Property Get Farbe_HalberTag() As OLE_COLOR
Attribute Farbe_HalberTag.VB_Description = "Gibt die Farbe zurück, die zum Auffüllen von Figuren, Kreisen und Feldern verwendet wird, oder legt diese fest."
    Farbe_HalberTag = m_Farbe_HalberTag
End Property

Public Property Let Farbe_HalberTag(ByVal New_Farbe_HalberTag As OLE_COLOR)
    m_Farbe_HalberTag = New_Farbe_HalberTag
    PropertyChanged "Farbe_HalberTag"
End Property

Public Property Get Farbe_einzelner_Tag() As OLE_COLOR
Attribute Farbe_einzelner_Tag.VB_Description = "Gibt die Farbe zurück, die zum Auffüllen von Figuren, Kreisen und Feldern verwendet wird, oder legt diese fest."
    Farbe_einzelner_Tag = m_Farbe_einzelner_Tag
End Property

Public Property Let Farbe_einzelner_Tag(ByVal New_Farbe_einzelner_Tag As OLE_COLOR)
    m_Farbe_einzelner_Tag = New_Farbe_einzelner_Tag
    PropertyChanged "Farbe_einzelner_Tag"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gibt die Farbe eines Objektrahmens zurück oder legt diese fest."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Legende() As Boolean
Attribute Legende.VB_ProcData.VB_Invoke_Property = "Zeitraum"
    Legende = m_Legende
End Property

Public Property Let Legende(ByVal New_Legende As Boolean)
    m_Legende = New_Legende
    PropertyChanged "Legende"
End Property

Public Property Get Filter_Feldname() As String
Attribute Filter_Feldname.VB_Description = "Optionaler Filter des Recordsets"
Attribute Filter_Feldname.VB_ProcData.VB_Invoke_Property = "Zeitraum"
    Filter_Feldname = m_Filter_Feldname
End Property

Public Property Let Filter_Feldname(ByVal New_Filter_Feldname As String)
    m_Filter_Feldname = New_Filter_Feldname
    PropertyChanged "Filter_Feldname"
End Property

Public Property Get ToolTip_Y() As String
Attribute ToolTip_Y.VB_ProcData.VB_Invoke_Property = "Zeitraum"
    ToolTip_Y = m_ToolTip_Y
End Property

Public Property Let ToolTip_Y(ByVal New_ToolTip_Y As String)
    m_ToolTip_Y = New_ToolTip_Y
    PropertyChanged "ToolTip_Y"
End Property

Public Property Get Filter_Bedingung() As String
Attribute Filter_Bedingung.VB_ProcData.VB_Invoke_Property = "Zeitraum"
    Filter_Bedingung = m_Filter_Bedingung
End Property

Public Property Let Filter_Bedingung(ByVal New_Filter_Bedingung As String)
    m_Filter_Bedingung = New_Filter_Bedingung
    PropertyChanged "Filter_Bedingung"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=Frame_ALL,Frame_ALL,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."
    BackColor = Frame_ALL.BackColor
End Property
'
'Public Property Get ConnectionString() As Variant
'    ConnectionString = m_ConnectionString
'End Property
'
'Public Property Let ConnectionString(ByVal New_ConnectionString As Variant)
'    m_ConnectionString = New_ConnectionString
'    PropertyChanged "ConnectionString"
'End Property
'
Public Property Get ConnectionString() As String
Attribute ConnectionString.VB_ProcData.VB_Invoke_Property = "Datenbank"
    ConnectionString = m_ConnectionString
End Property

Public Property Let ConnectionString(ByVal New_ConnectionString As String)
    m_ConnectionString = New_ConnectionString
    PropertyChanged "ConnectionString"
End Property

