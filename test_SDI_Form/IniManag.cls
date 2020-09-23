VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniManag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'La classe IniManag e' costituita
'da due funzioni 'operative' (Leggi e Scrivi) e
'da due coppie di funzioni Property per le due proprieta' (NomeFileIni e Sezione).
'La Property Let per il nome del file permette di passare anche il solo nome,
'nel qual caso viene aggiunta l'estensione di default ('ini') ed il percorso dell'applicazione.
'Le due funzioni Leggi e Scrivi fanno uso delle funzioni API dichiarate private nell'apposita sezione.
'N.B. La funzione Leggi restituisce sempre una stringa
'(il che significa fare eventuali conversioni se si vuole una data o un numero).

' =========================================================================

'Esempio d'uso per leggere una riga(indice) di un tipico file setup.lst:
'tipicamente, una riga della sezione Files definisce la chiave Filex (dove x e' l'indice) con una serie di informazioni
'(e' piu' facile andare a leggerne uno col Notepad per capire cio' che ho appena cercato di spiegare)

'Private Function LeggiIni(indice As Integer) As String
'Dim FileIni As New IniManag
'Dim riga As String

'With FileIni
' imposto il nome del file, che puo' anche essere una costante
' definita in altra parte del codice
'  .nome = "Setup.lst"
' imposto la sezione in cui cercare un valore
'  .sezione = "Files"
' passo alla funzione di lettura la chiave da leggere e l'eventuale risposta di default e ne ricevo il valore
'  LeggiIni = .leggi("File" & CStr(indice), "")
'End With
' elimino esplicitamente l'oggetto inutile prima di uscire
'Set FileIni = Nothing
'End Function

' =========================================================================

'Esempio d 'uso per scrivere una riga di un tipico file setup.lst:

'Private Function ScriviIni(riga As String, indice) As String
'Dim FileIni As New IniManag
'Dim ret As Long

'With FileIni
' imposto il nome del file, che puo' anche essere una costante definita in altra parte del codice
'  .nome = "Setup.lst"
' imposto la sezione in cui scrivere un valore, se non c'e' viene creata
'  .sezione = "Files"
' passo alla funzione di scrittura la chiave da trovare/creare ed il valore da scrivere
'ret = .scrivi("File" & CStr(indice), riga)
'End With
' elimino esplicitamente l'oggetto inutile prima di uscire
'Set FileIni = Nothing
'End Function

' =========================================================================

Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
  (ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpString As Any, _
  ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
  (ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpDefault As String, _
  ByVal lpReturnedString As String, _
  ByVal nSize As Long, _
  ByVal lpFileName As String) As Long

Private Const maxl As Integer = 255
Private Const z As Integer = 0

Private pNomeFileIni As String
Private pSezione As String

Public Property Let nome(ByVal valore As String)
  pNomeFileIni = valore
  If InStr(pNomeFileIni, ".") = z Then pNomeFileIni = pNomeFileIni & ".ini"
  If InStr(pNomeFileIni, "\") = z Then pNomeFileIni = App.Path & "\" & pNomeFileIni
End Property

Public Property Get nome() As String
  nome = pNomeFileIni
End Property

Public Property Get sezione() As String
  sezione = pSezione
End Property

Public Property Let sezione(ByVal valore As String)
  pSezione = valore
End Property

Public Function scrivi(chiave As String, valore As String) As Long
  scrivi = WritePrivateProfileString(pSezione, chiave, valore, pNomeFileIni)
End Function

Public Function leggi(chiave As String, default As String) As String
  Dim valore As String, lret As Long
  valore = String$(maxl, z)
  lret = GetPrivateProfileString(pSezione, chiave, default, valore, maxl, pNomeFileIni)
  leggi = Left(valore, lret)
End Function



