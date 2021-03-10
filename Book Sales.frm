VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   8076
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   12108
   LinkTopic       =   "Form1"
   ScaleHeight     =   8076
   ScaleWidth      =   12108
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C00000&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C00000&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C00000&
      Caption         =   "Clear Sale"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H00C00000&
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "Frame2"
      Height          =   2895
      Left            =   1080
      TabIndex        =   7
      Top             =   4200
      Width           =   8175
      Begin VB.Label lblDiscountedPrice 
         Height          =   495
         Left            =   2400
         TabIndex        =   18
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblDiscount 
         Height          =   495
         Left            =   2400
         TabIndex        =   17
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblExtendedPrice 
         Height          =   495
         Left            =   2400
         TabIndex        =   16
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C000&
         Caption         =   "Discounted Price"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "15% Discount"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "Extended Price"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.TextBox txtPrice 
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "R 'n R for Reading 'n Refreshment"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   3015
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   8175
      Begin VB.TextBox txtTitle 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtQuantity 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Book Sales"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3600
      TabIndex        =   15
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Book Sales
'Date: April 2017
'Programmer: Nasuha Asri
'Description: This project demonstrate the use of variables, contants, and calculations
'Folder: Asy.Haz

Option Explicit
Const mcurDiscountRate                     As Currency = 0.15

Private Sub cmdCalculate_Click()
        'Calculate the price and discount
        
        Dim intQuantity        As Integer
        Dim curPrice           As Currency
        Dim curExtendedPrice   As Currency
        Dim curDiscount        As Currency
        Dim curDiscountRate    As Currency
        Dim curDiscountedPrice As Currency
    
        
        'Convert input values to numeric variables
        intQuantity = Val(txtQuantity.Text)
        curPrice = Val(txtPrice.Text)
        
        'Calculate values
        curExtendedPrice = intQuantity * curPrice
        curDiscount = curExtendedPrice * mcurDiscountRate
        curDiscountedPrice = curExtendedPrice - curDiscount
        
        'Format and display answers
        lblExtendedPrice.Caption = FormatCurrency(curExtendedPrice)
        lblDiscount.Caption = FormatNumber(curDiscount, 2)
        lblDiscountedPrice.Caption = FormatCurrency(curDiscountedPrice)
        
End Sub

Private Sub cmdClear_Click()
        'Clear previous amounts from the form
        
        txtQuantity.Text = ""
        txtTitle.Text = ""
        txtPrice.Text = ""
        lblExtendedPrice.Caption = ""
        lblDiscount.Caption = ""
        lblDiscountedPrice.Caption = ""
        txtQuantity.SetFocus
End Sub

Private Sub cmdExit_Click()
        'Exit the project
        
        End
End Sub


Private Sub cmdPrint_Click()
        'PrintForm
        PrintForm
End Sub

