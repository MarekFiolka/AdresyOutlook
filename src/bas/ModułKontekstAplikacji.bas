Attribute VB_Name = "Modu³KontekstAplikacji"
'@Folder("Outlook")
Option Explicit

Private Type TMKontekstAplikacji
    AppContext As KontekstAplikacji
End Type

Private this As TMKontekstAplikacji

'@Description "Zwraca wspoldzielony kontekst aplikacji."
Public Property Get AppContext() As KontekstAplikacji
Attribute AppContext.VB_Description = "Zwraca wspoldzielony kontekst aplikacji."

    If this.AppContext Is Nothing Then
        Set this.AppContext = New KontekstAplikacji
    End If
    
    Set AppContext = this.AppContext

End Property

'@Description "Czysci wspoldzielony kontekst aplikacji."
'@EntryPoint
Public Sub ResetAppContext()
Attribute ResetAppContext.VB_Description = "Czysci wspoldzielony kontekst aplikacji."
    Set this.AppContext = Nothing
End Sub
