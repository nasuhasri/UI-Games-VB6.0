VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14415
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   Picture         =   "Games.frx":0000
   ScaleHeight     =   7920
   ScaleWidth      =   14415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H008080FF&
      Caption         =   "Add To Orders"
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox lblNBVG 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4320
      TabIndex        =   23
      Top             =   4080
      Width           =   2535
   End
   Begin VB.OptionButton OptPSVita 
      BackColor       =   &H00C0C0FF&
      Caption         =   "PlayStation Vita"
      Height          =   495
      Left            =   2160
      TabIndex        =   21
      Top             =   4080
      Width           =   1935
   End
   Begin VB.OptionButton OptXBox1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "XBox One"
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   4080
      Width           =   1575
   End
   Begin VB.OptionButton OptPS3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "PlayStation 3"
      Height          =   495
      Left            =   2160
      TabIndex        =   19
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H008080FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H008080FF&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H008080FF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H008080FF&
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox lblDiscountedPrice 
      BackColor       =   &H00C0E0FF&
      Height          =   405
      Left            =   6720
      TabIndex        =   13
      Top             =   7200
      Width           =   1935
   End
   Begin VB.TextBox lblDiscount 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox lblExtendedPrice 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   7320
      Width           =   1815
   End
   Begin VB.TextBox txtQuantity 
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   6600
      Width           =   1815
   End
   Begin VB.TextBox txtPrice 
      BackColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   2040
      TabIndex        =   5
      Top             =   5760
      Width           =   1815
   End
   Begin VB.ComboBox ComboVG 
      BackColor       =   &H0080C0FF&
      Height          =   315
      ItemData        =   "Games.frx":1806F
      Left            =   240
      List            =   "Games.frx":1808E
      TabIndex        =   3
      Top             =   5280
      Width           =   3855
   End
   Begin VB.OptionButton OptPS4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "PlayStation 4"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblOrder 
      BackStyle       =   0  'Transparent
      Caption         =   "Order"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   4680
      TabIndex        =   22
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label ChooseCategory 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose your category"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   18
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Image ImgCall 
      Height          =   2655
      Left            =   7080
      Picture         =   "Games.frx":1811F
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   3120
   End
   Begin VB.Image ImgSecond 
      Appearance      =   0  'Flat
      Height          =   2655
      Left            =   10320
      Picture         =   "Games.frx":23AE4
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Image ImgAssassins 
      Height          =   2355
      Left            =   240
      Picture         =   "Games.frx":3A754
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3600
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Discounted Price"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   4560
      TabIndex        =   12
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "15% Discount"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Price"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label NBVG 
      BackStyle       =   0  'Transparent
      Caption         =   "New Best Video Games"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   4800
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HazAsy World of Fantasy and Dreams"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1575
      Index           =   0
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: HazAsy World of Dream and Fantasy
'Date: April 2017
'Programmer: Nasuha Asri
'Description: The best new video games for the best gamers
'Folder: Asyhaz

Option Explicit
Const mcurDiscountRate                 As Currency = 0.15

Private Sub cmdAdd_Click()
lblNBVG.Text = ComboVG.Text
End Sub

Private Sub cmdCalculate_Click()
       'Calculate the price and discount
       
       Dim intQuantity            As Integer
       Dim curPrice               As Currency
       Dim curExtendedPrice       As Currency
       Dim curDiscount            As Currency
       Dim curDiscountedPrice     As Currency
       
       'Convert input values to numeric variables
       intQuantity = Val(txtQuantity.Text)
       curPrice = Val(txtPrice.Text)
       
       'Calculate values
       curExtendedPrice = intQuantity * curPrice
       curDiscount = curExtendedPrice * mcurDiscountRate
       curDiscountedPrice = curExtendedPrice - curDiscount
       
       'Format and display answers
       lblExtendedPrice.Text = FormatCurrency(curExtendedPrice)
       lblDiscount.Text = FormatNumber(curDiscount, 2)
       lblDiscountedPrice.Text = FormatCurrency(curDiscountedPrice)
       End Sub
       
Private Sub cmdClear_Click()
   'Clear previous amounts from the form
        
        txtQuantity.Text = ""
        txtPrice.Text = ""
        lblExtendedPrice.Text = ""
        lblDiscount.Text = ""
        lblDiscountedPrice.Text = ""
        txtQuantity.SetFocus
        lblNBVG.Text = ""
        ComboVG.Text = ""
End Sub

Private Sub cmdExit_Click()
'Exit the project
        
        End
End Sub

Private Sub cmdPrint_Click()
        'PrintForm
        PrintForm
End Sub

Private Sub ComboVG_Change()

End Sub





Private Sub OptPS3_Click()
If OptPS3.Enabled = True Then
ComboVG.Clear
ComboVG.AddItem "MLB 16:The Show"
ComboVG.AddItem "The Bit.Trip"
ComboVG.AddItem "Grand Theft Auto V"
ComboVG.AddItem "Dark Souls 2"
ComboVG.AddItem "King's Field II"
ComboVG.AddItem "Watch Dogs"
ComboVG.AddItem "The Last of Us"
ComboVG.AddItem "Uncharted 2: Among Thieves"
ComboVG.AddItem "Fifa 16"
ComboVG.AddItem "Call of Duty: Modern Warfare 3"
End If
End Sub

Private Sub OptPS4_Click()
If OptPS4.Enabled = True Then
ComboVG.Clear
ComboVG.AddItem "Uncharted 4"
ComboVG.AddItem "Final Fantasy XV"
ComboVG.AddItem "The Last Guardian"
ComboVG.AddItem "NieRAutomata"
ComboVG.AddItem "Fifa 17"
ComboVG.AddItem "Battlefield 1"
ComboVG.AddItem "Bloodborne"
ComboVG.AddItem "TitanFall 2"
ComboVG.AddItem "The Witcher 3: Wild Hunt"
End If
End Sub

Private Sub OptPSVita_Click()
If OptPSVita.Enabled = True Then
ComboVG.Clear
ComboVG.AddItem "Borderlands 2"
ComboVG.AddItem "Conception II: Children of the Seven Stars"
ComboVG.AddItem "CounterSpy"
ComboVG.AddItem "Demon Gaze"
ComboVG.AddItem "Danganronpa: Trigger Happy Havoc"
ComboVG.AddItem "Destiny of Spirits"
ComboVG.AddItem "Dragon Ball Z: The Battle of Z"
ComboVG.AddItem "Dustforce"
ComboVG.AddItem "Final Fantasy X|X-2 HD"
ComboVG.AddItem "Freedom Wars"
End If
End Sub

Private Sub OptXBox1_Click()
If OptXBox1.Enabled = True Then
ComboVG.Clear
ComboVG.AddItem "Rise of The Tomb Rider"
ComboVG.AddItem "Grand Theft Auto V"
ComboVG.AddItem "Hitman"
ComboVG.AddItem "Inside"
ComboVG.AddItem "Tittanfall 2"
ComboVG.AddItem "Forza Horizon 3"
ComboVG.AddItem "The Witcher 3"
ComboVG.AddItem "Dark Souls III"
ComboVG.AddItem "Destiny"
ComboVG.AddItem "Overwatch"
End If
End Sub



