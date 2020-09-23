VERSION 5.00
Begin VB.Form CSV2TXT_Main 
   Caption         =   "Kevin Ritch - V8Software.com"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   Picture         =   "CSV2TXT.frx":0000
   ScaleHeight     =   4575
   ScaleWidth      =   12045
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "CSV2TXT_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NulString As String
Dim Q As String
Dim C As String
Dim tb As String
Dim QCQ As String
Dim StoreComma As String
Dim DummyQuote As String
Dim RAW
Private Sub Form_Load()
 tb = Chr$(9)
 Q = Chr$(34)
 C = ","
 QCQ = Q & C & Q
 StoreComma = Chr$(255)
 DummyQuote = Chr$(254)
'=================================================================================
'CREATE YOUR "Occasional Quotes File" (Commas in Street Addresses etc) File FIRST!
'=================================================================================
 CSVFile$ = "c:\V8Client\OccasionalQuotes.csv"
 TabFile$ = "c:\V8Client\ImportOK.tab"
 IsQuoteCommaQuoteFile = TestQCQ(CSVFile$)
 Open CSVFile$ For Input As #1
 Open TabFile$ For Output As #2
 While Not EOF(1)
  Line Input #1, a$
  If IsQuoteCommaQuoteFile Then
   a$ = Replace$(a$, QCQ, tb)
   a$ = Replace$(a$, Q, NulString)
  Else ' OK This file is not a fully fledged Quote Comma Quote file
  '=========================== ===================
  'Find Commas in WITHIN cells (ie NOT Delimiters)
  '=========================== ===================
   While InStr(a$, Q)
    S1 = InStr(a$, Q)
    S2 = InStr(S1 + 1, a$, Q)
    tmp$ = Mid$(a$, S1 + 1, (S2 - S1) - 1)
   '=========================   =============================
   'EG "166, Riviera Parkway" - TEMPORARILY "STORE" THE COMMA
   '=========================   =============================
    tmp$ = Replace(tmp$, C, StoreComma)
    Mid$(a$, S1, 1) = DummyQuote
    Mid$(a$, S2, 1) = DummyQuote
    Mid$(a$, S1 + 1, (S2 - S1) - 1) = tmp$
   Wend
  '======================
  'Remove the DummyQuotes
  '======================
   a$ = Replace$(a$, DummyQuote, NulString)
  '========================
  'REPLACE COMMAS WITH TABS
  '========================
   a$ = Replace$(a$, C, tb)
  '=============================
  'Now replace the stored commas
  '=============================
   a$ = Replace$(a$, StoreComma, C)
  End If
 '=============================
 'SAVE THE LINE TO THE TAB FILE
 '=============================
  Print #2, a$
 Wend
 Close
 MsgBox "FILE CONVERTED"
 Shell "Notepad " & TabFile$, vbMaximizedFocus
 End
End Sub
Function TestQCQ(Filename As String) As Boolean
 TestQCQ = True
 DF = FreeFile
 Open Filename$ For Input As #DF
 While Not EOF(1)
  Line Input #DF, a$
  If InStr(a$, QCQ) = 0 Then
   TestQCQ = False
   GoTo DoneChecking:
  End If
 Wend
DoneChecking:
 Close #DF
End Function
