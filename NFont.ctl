VERSION 5.00
Begin VB.UserControl NFont 
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1485
   ScaleHeight     =   705
   ScaleWidth      =   1485
   ToolboxBitmap   =   "NFont.ctx":0000
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NFont"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NFont"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   30
      TabIndex        =   1
      Top             =   45
      Width           =   1410
   End
End
Attribute VB_Name = "NFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Property Let BackColor(bc As OLE_COLOR)
UserControl.BackColor = bc
End Property
Public Property Get BackColor() As OLE_COLOR
BackColor = UserControl.BackColor
End Property

Public Property Let ForeColor(fc As OLE_COLOR)
Label1.ForeColor = fc
End Property
Public Property Get ForeColor() As OLE_COLOR
ForeColor = Label1.ForeColor
End Property

Public Property Let Fontname(fo As String)
Label1.Fontname = fo
Label2.Fontname = fo
UserControl.Height = Label1.Height + 10
UserControl.Width = Label1.Width
End Property
Public Property Get Fontname() As String
Fontname = Label1.Fontname
End Property

Public Property Let Fontbold(Fontbold As Boolean)
Label1.Fontbold = Fontbold
Label2.Fontbold = Fontbold
UserControl.Height = Label1.Height + 10
UserControl.Width = Label1.Width
End Property
Public Property Get Fontbold() As Boolean
Fontbold = Label1.Fontbold
End Property

Public Property Let Fontsize(Fontsize As String)
Label1.Fontsize = Fontsize
Label2.Fontsize = Fontsize
UserControl.Height = Label1.Height + 10
UserControl.Width = Label1.Width
End Property
Public Property Get Fontsize() As String
Fontsize = Label1.Fontsize
End Property

Public Property Let FontUnderline(FontUnderline As Boolean)
Label1.FontUnderline = FontUnderline
Label2.FontUnderline = FontUnderline
UserControl.Height = Label1.Height + 10
UserControl.Width = Label1.Width
End Property
Public Property Get FontUnderline() As Boolean
FontUnderline = Label1.FontUnderline
End Property


Public Property Let Caption(ca As String)
Label1.Caption = ca
Label2.Caption = ca
UserControl.Height = Label1.Height + 10
UserControl.Width = Label1.Width
End Property
Public Property Get Caption() As String
Caption = Label1.Caption
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Label1.Caption = PropBag.ReadProperty("CA")
Label2.Caption = PropBag.ReadProperty("CA")
UserControl.BackColor = PropBag.ReadProperty("BC")
Label1.ForeColor = PropBag.ReadProperty("FC")
Label1.Fontbold = PropBag.ReadProperty("fb")
Label2.Fontbold = PropBag.ReadProperty("fb")
Label1.Fontsize = PropBag.ReadProperty("fs")
Label2.Fontsize = PropBag.ReadProperty("fs")
Label1.FontUnderline = PropBag.ReadProperty("fu")
Label2.FontUnderline = PropBag.ReadProperty("fu")
Label1.Fontname = PropBag.ReadProperty("fn")
Label2.Fontname = PropBag.ReadProperty("fn")

End Sub

Private Sub UserControl_Resize()
'UserControl.Height = Label1.Height + 10
'UserControl.Width = Label1.Width
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "CA", Label1.Caption
PropBag.WriteProperty "BC", UserControl.BackColor
PropBag.WriteProperty "FC", Label1.ForeColor
PropBag.WriteProperty "fb", Label1.Fontbold
PropBag.WriteProperty "fs", Label1.Fontsize
PropBag.WriteProperty "fu", Label1.FontUnderline
PropBag.WriteProperty "fn", Label1.Fontname

End Sub
