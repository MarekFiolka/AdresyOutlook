VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KodPocztowy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule AnnotationInIncompatibleComponentType
'@Folder Kody
Option Explicit

Private Type TKodPocztowy
    Kod As String
    Adres As String
    Miejscowosc As String
    Wojewodztwo As String
    Powiat As String
End Type

Private this As TKodPocztowy

'@Description("Kod pocztowy")
Public Property Get Kod() As String
Attribute Kod.VB_Description = "Kod pocztowy"
    Kod = this.Kod
End Property

'@Description("Ustawia kod pocztowy")
'@EntryPoint
Public Property Let Kod(ByVal Value As String)
Attribute Kod.VB_Description = "Ustawia kod pocztowy"
    this.Kod = Trim$(Value)
End Property

'@Description("Adres")
'@EntryPoint
Public Property Get Adres() As String
Attribute Adres.VB_Description = "Adres"
    Adres = this.Adres
End Property

'@Description("Ustawia adres")
'@EntryPoint
Public Property Let Adres(ByVal Value As String)
Attribute Adres.VB_Description = "Ustawia adres"
    this.Adres = Trim$(Value)
End Property

'@Description("Miejscowoœæ")
Public Property Get Miejscowosc() As String
Attribute Miejscowosc.VB_Description = "Miejscowoœæ"
    Miejscowosc = this.Miejscowosc
End Property

'@Description("Ustawia miejscowoœæ")
'@EntryPoint
Public Property Let Miejscowosc(ByVal Value As String)
Attribute Miejscowosc.VB_Description = "Ustawia miejscowoœæ"
    this.Miejscowosc = Trim$(Value)
End Property

'@Description("Województwo")
Public Property Get Wojewodztwo() As String
Attribute Wojewodztwo.VB_Description = "Województwo"
    Wojewodztwo = this.Wojewodztwo
End Property

'@Description("Ustawia województwo")
'@EntryPoint
Public Property Let Wojewodztwo(ByVal Value As String)
Attribute Wojewodztwo.VB_Description = "Ustawia województwo"
    this.Wojewodztwo = Trim$(Value)
End Property

'@Description("Powiat")
'@EntryPoint
Public Property Get Powiat() As String
Attribute Powiat.VB_Description = "Powiat"
    Powiat = this.Powiat
End Property

'@Description("Ustawia powiat")
'@EntryPoint
Public Property Let Powiat(ByVal Value As String)
Attribute Powiat.VB_Description = "Ustawia powiat"
    this.Powiat = Trim$(Value)
End Property

'@Description("Inicjalizuje rekord kodu pocztowego")
Public Sub Init( _
    ByVal Kod As String, _
    ByVal Adres As String, _
    ByVal Miejscowosc As String, _
    ByVal Wojewodztwo As String, _
    ByVal Powiat As String _
)
Attribute Init.VB_Description = "Inicjalizuje rekord kodu pocztowego"
    this.Kod = Trim$(Kod)
    this.Adres = Trim$(Adres)
    this.Miejscowosc = Trim$(Miejscowosc)
    this.Wojewodztwo = Trim$(Wojewodztwo)
    this.Powiat = Trim$(Powiat)
End Sub

'@Description("Zwraca reprezentacjê tekstow¹ obiektu")
'@EntryPoint
Public Function ToString() As String
Attribute ToString.VB_Description = "Zwraca reprezentacjê tekstow¹ obiektu"

    ToString = _
        "Kod: " & this.Kod & vbCrLf & _
        "Adres: " & this.Adres & vbCrLf & _
        "Miejscowoœæ: " & this.Miejscowosc & vbCrLf & _
        "Województwo: " & this.Wojewodztwo & vbCrLf & _
        "Powiat: " & this.Powiat

End Function

'@Description("Zwraca jednowierszow¹ reprezentacjê obiektu")
'@EntryPoint
Public Function ToDebugString() As String
Attribute ToDebugString.VB_Description = "Zwraca jednowierszow¹ reprezentacjê obiektu"

    ToDebugString = _
        "[" & this.Kod & "] " & _
        this.Miejscowosc & ", " & _
        this.Wojewodztwo & " | " & _
        this.Adres

End Function
