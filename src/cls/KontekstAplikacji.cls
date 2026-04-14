VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KontekstAplikacji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule AnnotationInIncompatibleComponentType
'@Folder("Outlook")
Option Explicit

Private Type TKontekstAplikacji
    KodyPocztowe As KodyPocztowe
    SciezkaKodowPocztowych As String
    CzyKodyPocztoweZainicjalizowane As Boolean
End Type

Private this As TKontekstAplikacji

Private Sub Class_Initialize()
    #If DEV Then
        this.SciezkaKodowPocztowych = "E:\VBA\AdresyOutlook\docs\Kody pocztowe.csv"
    #Else
        this.SciezkaKodowPocztowych = "C:\Users\marek\Documents\Pliki programu Outlook\Kody pocztowe.csv"
    #End If
    this.SciezkaKodowPocztowych = "C:\Users\marek\Documents\Pliki programu Outlook\Kody pocztowe.csv"
    this.CzyKodyPocztoweZainicjalizowane = False
End Sub

Private Sub Class_Terminate()
    Set this.KodyPocztowe = Nothing
End Sub

'@Description("Œcie¿ka do pliku CSV z baz¹ kodów pocztowych.")
'@EntryPoint
Public Property Get SciezkaKodowPocztowych() As String
Attribute SciezkaKodowPocztowych.VB_Description = "Œcie¿ka do pliku CSV z baz¹ kodów pocztowych."
    SciezkaKodowPocztowych = this.SciezkaKodowPocztowych
End Property

'@Description("Ustawia œcie¿kê do pliku CSV z baz¹ kodów pocztowych.")
'@EntryPoint
Public Property Let SciezkaKodowPocztowych(ByVal Value As String)
Attribute SciezkaKodowPocztowych.VB_Description = "Ustawia œcie¿kê do pliku CSV z baz¹ kodów pocztowych."

    Dim localValue As String
    localValue = Value
    localValue = Trim$(localValue)
    
    If Len(localValue) = 0 Then
        Err.Raise vbObjectError + 7000, TypeName(Me), _
                  "SciezkaKodowPocztowych nie mo¿e byæ pusta."
    End If
    
    If StrComp(this.SciezkaKodowPocztowych, localValue, vbBinaryCompare) <> 0 Then
        this.SciezkaKodowPocztowych = localValue
        ResetKodyPocztowe
    End If

End Property

'@Description("Zwraca wspó³dzielon¹ bazê kodów pocztowych z kontrol¹ inicjalizacji.")
Public Property Get KodyPocztowe() As KodyPocztowe
Attribute KodyPocztowe.VB_Description = "Zwraca wspó³dzielon¹ bazê kodów pocztowych z kontrol¹ inicjalizacji."
    EnsureKodyPocztowe
    Set KodyPocztowe = this.KodyPocztowe
End Property

'@Description("Wymusza ponowne wczytanie bazy kodów pocztowych.")
'@EntryPoint
Public Sub ReloadKodyPocztowe()
Attribute ReloadKodyPocztowe.VB_Description = "Wymusza ponowne wczytanie bazy kodów pocztowych."
    ResetKodyPocztowe
    EnsureKodyPocztowe
End Sub

'@Description("Czyœci stan bazy kodów pocztowych. Kolejny dostêp spowoduje ponown¹ inicjalizacjê.")
Public Sub ResetKodyPocztowe()
Attribute ResetKodyPocztowe.VB_Description = "Czyœci stan bazy kodów pocztowych. Kolejny dostêp spowoduje ponown¹ inicjalizacjê."
    this.CzyKodyPocztoweZainicjalizowane = False
    Set this.KodyPocztowe = Nothing
End Sub

'@Description("Zwraca województwo dla wskazanego kodu pocztowego.")
Public Function WojewodztwoDlaKodu(ByVal Kod As String) As String
Attribute WojewodztwoDlaKodu.VB_Description = "Zwraca województwo dla wskazanego kodu pocztowego."
    EnsureKodyPocztowe
    WojewodztwoDlaKodu = this.KodyPocztowe.WojewodztwoDlaKodu(Kod)
End Function

'@Description("Zwraca miejscowoœæ dla wskazanego kodu pocztowego.")
Public Function MiejscowoscDlaKodu(ByVal Kod As String) As String
Attribute MiejscowoscDlaKodu.VB_Description = "Zwraca miejscowoœæ dla wskazanego kodu pocztowego."
    EnsureKodyPocztowe
    MiejscowoscDlaKodu = this.KodyPocztowe.MiejscowoscDlaKodu(Kod)
End Function

'@Description("Zwraca True, gdy baza kodów pocztowych zosta³a ju¿ zainicjalizowana.")
'@EntryPoint
Public Property Get CzyKodyPocztoweZainicjalizowane() As Boolean
Attribute CzyKodyPocztoweZainicjalizowane.VB_Description = "Zwraca True, gdy baza kodów pocztowych zosta³a ju¿ zainicjalizowana."
    CzyKodyPocztoweZainicjalizowane = this.CzyKodyPocztoweZainicjalizowane
End Property

'@Description("Zwraca krótki opis stanu kontekstu aplikacji.")
Public Function ToDebugString() As String
Attribute ToDebugString.VB_Description = "Zwraca krótki opis stanu kontekstu aplikacji."

    Dim LiczbaKodow As Long
    
    If this.CzyKodyPocztoweZainicjalizowane Then
        If Not this.KodyPocztowe Is Nothing Then
            LiczbaKodow = this.KodyPocztowe.Count
        End If
    End If
    
    ToDebugString = _
        "SciezkaKodowPocztowych=""" & this.SciezkaKodowPocztowych & """" & _
        " | CzyKodyPocztoweZainicjalizowane=" & CStr(this.CzyKodyPocztoweZainicjalizowane) & _
        " | LiczbaKodow=" & CStr(LiczbaKodow)

End Function

Private Sub EnsureKodyPocztowe()

    If this.CzyKodyPocztoweZainicjalizowane Then
        Exit Sub
    End If
    
    ValidateSciezkaKodowPocztowych this.SciezkaKodowPocztowych
    
    Set this.KodyPocztowe = New KodyPocztowe
    this.KodyPocztowe.LoadFromCsv this.SciezkaKodowPocztowych
    this.CzyKodyPocztoweZainicjalizowane = True

End Sub

Private Sub ValidateSciezkaKodowPocztowych(ByVal Value As String)

    If Len(Trim$(Value)) = 0 Then
        Err.Raise vbObjectError + 7001, TypeName(Me), _
                  "Nie ustawiono œcie¿ki do pliku z kodami pocztowymi."
    End If
    
    If Dir$(Value) = vbNullString Then
        Err.Raise vbObjectError + 7002, TypeName(Me), _
                  "Nie znaleziono pliku: " & Value
    End If

End Sub
