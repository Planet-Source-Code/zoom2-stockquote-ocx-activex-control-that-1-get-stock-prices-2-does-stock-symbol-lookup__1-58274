VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.UserControl stockQuote 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   435
   InvisibleAtRuntime=   -1  'True
   Picture         =   "stockQuote.ctx":0000
   ScaleHeight     =   375
   ScaleWidth      =   435
   ToolboxBitmap   =   "stockQuote.ctx":0102
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   240
      Left            =   540
      TabIndex        =   0
      Top             =   585
      Width           =   240
      ExtentX         =   423
      ExtentY         =   423
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "stockQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum enStockType
   [Stocks] = 1
   [ETF] = 2
   [Indices] = 3
   [Mutual Funds] = 4
   [Options] = 5
End Enum

Enum enMarket
    [US] = 1
    [Worldwide] = 2
End Enum

Dim bRetrieved As Boolean
Dim bIsQuery     As Boolean
Dim bIsLookup  As Boolean

Event InvalidStockSymbol()
Event StockSymbolLookup(strStockSymbol As String)
Event StockSymbolLookupComplete()
Event priceDownloadProgress(strProg As String)
Event StockPrice(strPrice As String, currPrice As Currency)
 
'||||||||||||||| USER DOESNT KNOW THE STOCKS SYMBOL ||||||||||||||||||||
Sub symbol_lookup(sStockQueryString As String, _
                 stockType As enStockType, _
                 Market As enMarket)
Dim sQuery  As String, sType As String, sMarket As String
On Error GoTo local_error:

  bIsQuery = True
  bIsLookup = False
  bRetrieved = False
  '-- replace any spaces with "+"
  sStockQueryString = Replace(sStockQueryString, " ", "+")
  sType = Choose(stockType, "S", "E", "I", "M", "O")
  sMarket = Choose(Market, "US", "")
  sQuery = "http://finance.yahoo.com/l?s=" & _
           sStockQueryString & "&t=" & sType & "&m=" & sMarket
  WebBrowser1.Navigate sQuery
  
Exit Sub
local_error:
   If Err.Number <> 0 Then
       Debug.Print "stockQuote.symbol_lookup: " & Err.Number & "." & Err.Description
       Err.Clear
       Resume Next
   End If
End Sub
'||||| RETRIEVES STOCK QUOTE FOR A STOCK |||||||||||||||||||||||||||||||
Sub get_stock_quote(sStockSymbol As String)
On Error GoTo local_error

   bIsQuery = False
   bIsLookup = True
   bRetrieved = False
   WebBrowser1.Navigate "http://finance.yahoo.com/q?s=" & sStockSymbol
    
Exit Sub
local_error:
   If Err.Number <> 0 Then
       Debug.Print "stockQuote.get_stock_quote: " & Err.Number & "." & Err.Description
       Err.Clear
       Resume Next
   End If
End Sub
'|||||||||||||| SHUT THE CHATTERBOX BROWSER UP |||||||||||||||||||||||||
Private Sub UserControl_Show()
On Error Resume Next

  With WebBrowser1
     .AddressBar = False
     .FullScreen = False
     .MenuBar = False
     .RegisterAsBrowser = False
     .RegisterAsDropTarget = False
     .Silent = True
     .StatusBar = False
     .Navigate "about:blank"
  End With
End Sub
'||||||| PROGRESS FEEDBACK ||||||||||||||||||||||||||||||||||||||
Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
   RaiseEvent priceDownloadProgress("Downloading stock price")
End Sub
Private Sub WebBrowser1_DownloadBegin()
   RaiseEvent priceDownloadProgress("Attempting to retrieve stock quote")
End Sub
 

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Dim cDoc          As HTMLDocument
Dim cGen          As HTMLGenericElement
Dim TBLtag        As HTMLTable
Dim TDtag         As IHTMLTableCell
Dim TBLtag2       As IHTMLTable2
Dim sParts1()     As String
Dim sParts2()     As String
Dim hrefStr       As String
Dim symbolRetStr  As String
Dim strTag        As String
Dim newStr        As String
Dim lcnt          As Long
Dim upper         As Long
Dim stock_quote   As String
On Error GoTo local_error:
 
 If WebBrowser1.LocationURL = "http:///" Or _
        WebBrowser1.LocationURL = "about:blank" Then Exit Sub
        
  Set cDoc = WebBrowser1.Document
  
  If Not (bRetrieved) Then '-- force return of results just once
      bRetrieved = True
1     RaiseEvent priceDownloadProgress("Download complete.  Parsing data")
     
     '-- bQuery means the user is trying to find a stock symbol for a company
8    If bIsQuery Then
        '-- retrieve the number of href tags
9       upper = (cDoc.getElementsByTagName("A").length - 1)
 
        '-- loop through each
        For lcnt = 0 To upper
            hrefStr = cDoc.getElementsByTagName("A")(lcnt).href
            symbolRetStr = StringExtract(hrefStr, "http://finance.yahoo.com/q?s=", "&d=t")
            If Len(Trim(symbolRetStr)) > 0 Then
                RaiseEvent StockSymbolLookup(symbolRetStr)
            End If
        Next lcnt
        RaiseEvent StockSymbolLookupComplete
        Exit Sub
     End If
' http://finance.yahoo.com/q?s=WMT&d=t
 
     If bIsLookup Then '-- were getting a stock quote..ticker symbol provided
        '-- look for the string ("net asset value:")
        '   the inner text immediately following is the price
3        stock_quote = Split(LCase(cDoc.body.innerText), "net asset value:")(1)
        '-- sometimes different formats on webpage for different type of
        '   stocks..ie mutual funds vs stocks, nasdaq vs nyse
4        If Len(Trim(stock_quote)) = 0 Then
5            stock_quote = Split(LCase(cDoc.body.innerText), "last trade:")(1)
         End If
         
6        stock_quote = Split(stock_quote, vbCrLf)(0)
        '--  pass price to event as both string and currency
7        RaiseEvent StockPrice(stock_quote, CCur(stock_quote))
      End If
      
 
      WebBrowser1.Stop '-- save on processor slightly
      Set cDoc = Nothing
  End If
  
Exit Sub
local_error:
       If Err.Number = 0 Then Exit Sub
  
       If Err.Number <> 9 Then
           RaiseEvent priceDownloadProgress("Error: " & Err.Description)
           Debug.Print "stockQuote.ocx.WebBrowser1_DocumentComplete: " & _
           Err.Description & "  line #: " & Erl()
           Err.Clear
           Resume Next
       Else '-- if we get "Subscript out of range" error on line 6 then
       '        its a sure sign an invalid ticker symbol was entered
           If Erl() = 6 Then
               RaiseEvent InvalidStockSymbol
               Exit Sub
           End If
       End If
       
       Resume Next
End Sub
 
'Extract a string from withing another string
Private Function StringExtract(ByVal baseString$, startStrTag$, _
                        Optional endStrTag$, _
                        Optional lngStartChrBuffer&, _
                        Optional lngEndChrBuffer&, _
                        Optional CaseMatters As Boolean = False) As String
On Error GoTo local_error:
Dim BS$, SST$, EST$, sTemp$
Dim SCB&, ECB&, iRet&, iRet2&
               '--==|| assign variables to shorter variables ||==--
               '--==|| for convenience,+ make all strings    ||==--
               '--==|| lcase if case doesnt matter
                BS$ = baseString
                If Not (CaseMatters) Then BS = LCase(BS)
                SST$ = startStrTag
                If Not (CaseMatters) Then SST = LCase(startStrTag)
                EST$ = endStrTag
                If Not (CaseMatters) Then EST = LCase(endStrTag)
                
                '--==|| assign val to SCB ||==--
                SCB = lngStartChrBuffer
                
                '--==|| assign val to ECB ||==--
                ECB = lngEndChrBuffer
                
                '--==|| search for the startStringTag ||==--
                iRet = InStr(1, BS, SST)
                If iRet <> 0 Then
                  'add the postion of the found startStringTag
                  'with the length of startStringTag and SCB
                  iRet = (iRet + Len(SST) + SCB)
                  'sTemp now holds from point iret to end
                  sTemp = Mid(BS, iRet, Len(BS) - (iRet - 1))
                  'exit function if EndStringTag not provided
                  If Len(Trim(EST)) <= 0 Then
                     StringExtract = sTemp
                     Exit Function
                  End If
                  iRet2 = InStr(1, sTemp, EST)
                  If iRet2 <> 0 Then
                     'add endChrBuffer to iRet2
                     iRet2 = (iRet2 + ECB)
                     StringExtract = Mid(sTemp, 1, (iRet2 - 1))
                  End If
                End If
local_error:
   If Err.Number <> 0 Then
       Debug.Print "stockQuote.StringExtract: " & Err.Number & "." & Err.Description
       Err.Clear
       Resume Next
   End If
End Function
 
 
 
 
 
 
 




' __ _     .     . _ __  _    . _     .    .
'/ _` | _ _  __ _ | '_ \| |__  (_) ___  _ _'
' (_| || '_\/ _` || |_) |  _ \ | |/ __|/ __|
'\__, || |   (_| || .__/| | | || | (__ \__ \
'|___/ |_|  \__,_||_|   |_| |_||_|\___||___/

Private Sub UserControl_Paint()
   UserControl.Line (0, 0)-(Width, Height), RGB(255, 255, 255), B
   UserControl.Line (-50, -50)-(Width - 20, Height - 20), RGB(180, 180, 190), B
End Sub
Private Sub UserControl_Resize()
           UserControl.Size _
  (16 * Screen.TwipsPerPixelX), (16 * Screen.TwipsPerPixelY)
End Sub

 

 
