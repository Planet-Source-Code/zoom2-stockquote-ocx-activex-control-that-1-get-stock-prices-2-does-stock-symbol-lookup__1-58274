VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stock Stop"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4455
      TabIndex        =   8
      Text            =   "Combo2"
      Top             =   270
      Width           =   1005
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3420
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   270
      Width           =   1050
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   3510
      TabIndex        =   5
      Top             =   630
      Width           =   1950
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1935
      TabIndex        =   4
      Text            =   "coca cola"
      Top             =   270
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Lookup stock symbol"
      Height          =   465
      Left            =   270
      TabIndex        =   3
      Top             =   90
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1980
      TabIndex        =   1
      ToolTipText     =   "try entering an invalid ticker symbol like ""afsafsdaf"" and watch you get slapped"
      Top             =   810
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Get stock quote"
      Height          =   465
      Left            =   270
      TabIndex        =   0
      Top             =   720
      Width           =   1635
   End
   Begin ProjStockQuoteControl.stockQuote stockQuote1 
      Left            =   5625
      Top             =   1080
      _extentx        =   423
      _extenty        =   423
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Height          =   915
      Left            =   180
      Top             =   675
      Width           =   3030
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Name of company    Stock Type         Market"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1935
      TabIndex        =   7
      Top             =   45
      Width           =   3525
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   285
      Left            =   270
      TabIndex        =   2
      Top             =   1260
      Width           =   2670
   End
   Begin VB.Menu mnuPop 
      Caption         =   "pop"
      Visible         =   0   'False
      Begin VB.Menu mnuClear 
         Caption         =   "Clear contents"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim m_stockType As Long, m_market As Long
 
'-- selecting the trade type [long]
Private Sub Combo1_Click()
   m_stockType = (Combo1.ListIndex + 1)
End Sub
'-- selecting the market [long]
Private Sub Combo2_Click()
  m_market = (Combo2.ListIndex + 1)
End Sub
'-- retrieve a stock price
Private Sub Command1_Click()
   stockQuote1.get_stock_quote Text1
End Sub
'-- look up a stock symbol for a company
Private Sub Command2_Click()
  stockQuote1.symbol_lookup Text2, m_stockType, m_market
End Sub
Private Sub Form_Load()
'-- add items to combo1 and combo2 and preselect the first item
  Combo1.AddItem "Stocks"
  Combo1.AddItem "ETF"
  Combo1.AddItem "Indices"
  Combo1.AddItem "Mutual Funds"
  Combo1.AddItem "Options"
  Combo1.ListIndex = 0
  Combo2.AddItem "US"
  Combo2.AddItem "Worldwide"
  Combo2.ListIndex = 0
End Sub
'-- place the stock symbol clicked on in list1 into the textbox
Private Sub List1_Click()
  Text1 = List1.List(List1.ListIndex)
End Sub
'-- show popup on list1_mousedown
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   PopupMenu mnuPop, , (List1.Left + List1.Width)
End Sub
'--clear listbox
Private Sub mnuClear_Click()
  List1.Clear
End Sub
'-- the quote for a stock was attempted to be retrieved
'-- and there was no valid return
Private Sub stockQuote1_InvalidStockSymbol()
   Label1 = "It appears " & Text1 & " is an invalid stock symbol"
End Sub
'-- report of the download progress of the control
Private Sub stockQuote1_priceDownloadProgress(strProg As String)
   Label1 = strProg
End Sub
'-- the price for the stock specified in text1 has been returned
Private Sub stockQuote1_StockPrice(strPrice As String, currPrice As Currency)
  Label1 = "$ " & strPrice
End Sub
'-- when looking up a stock symbol..when one or more are found
'-- this event is raised
Private Sub stockQuote1_StockSymbolLookup(strStockSymbol As String)
  List1.AddItem strStockSymbol
End Sub
'-- done searching for stock symbols
Private Sub stockQuote1_StockSymbolLookupComplete()
   MsgBox "Done with stock symbol lookup"
End Sub
