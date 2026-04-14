Attribute VB_Name = "NaprawiaczAdresu"
'@Folder("Adresy")
Option Explicit

Private Type TAdresRozpoznany
    Ulica As String
    Miasto As String
    Wojewodztwo As String
    KodPocztowy As String
    Kraj As String
    CzyPolska As Boolean
    Zrodlo As String
End Type

'@Description("Naprawia adres biznesowy aktualnie otwartego kontaktu Outlook.")
'@EntryPoint
Public Sub NaprawAdresBiznesowyBiezacegoKontaktu()
Attribute NaprawAdresBiznesowyBiezacegoKontaktu.VB_Description = "Naprawia adres biznesowy aktualnie otwartego kontaktu Outlook."

    Dim CurrentItem As Object
    Dim Contact As Outlook.ContactItem

    On Error Resume Next
    Set CurrentItem = Application.ActiveInspector.CurrentItem
    On Error GoTo 0

    If CurrentItem Is Nothing Then
        MsgBox "Nie ma otwartego elementu.", vbExclamation
        Exit Sub
    End If

    If TypeName(CurrentItem) <> "ContactItem" Then
        MsgBox "Aktualnie otwarty element nie jest kontaktem Outlook.", vbExclamation
        Exit Sub
    End If

    Set Contact = CurrentItem

    NaprawAdresBiznesowy Contact

End Sub

'@Description("Naprawia adres biznesowy wskazanego kontaktu Outlook.")
Public Sub NaprawAdresBiznesowy(ByVal Contact As Outlook.ContactItem)
Attribute NaprawAdresBiznesowy.VB_Description = "Naprawia adres biznesowy wskazanego kontaktu Outlook."

    Dim Parsed As TAdresRozpoznany
    Dim OldStreet As String
    Dim OldCity As String
    Dim OldState As String
    Dim OldPostalCode As String
    Dim OldCountry As String
    Dim Summary As String

    If Contact Is Nothing Then
        Err.Raise vbObjectError + 8100, TypeName(Contact), "Parametr Contact nie mo¿e byæ Nothing."
    End If

    Parsed = RozpoznajAdresBiznesowy(Contact)

    If Len(Parsed.KodPocztowy) = 0 Then
        MsgBox _
            "Nie uda³o siê rozpoznaæ kodu pocztowego." & vbCrLf & vbCrLf & _
            "Tekst Ÿród³owy:" & vbCrLf & Parsed.Zrodlo, _
            vbExclamation
        Exit Sub
    End If

    If Len(Parsed.Ulica) = 0 Then
        MsgBox _
            "Nie uda³o siê rozpoznaæ ulicy." & vbCrLf & vbCrLf & _
            "Tekst Ÿród³owy:" & vbCrLf & Parsed.Zrodlo, _
            vbExclamation
        Exit Sub
    End If

    If Len(Parsed.Miasto) = 0 Then
        MsgBox _
            "Nie uda³o siê rozpoznaæ miejscowoœci." & vbCrLf & vbCrLf & _
            "Tekst Ÿród³owy:" & vbCrLf & Parsed.Zrodlo, _
            vbExclamation
        Exit Sub
    End If

    If Parsed.CzyPolska Then
        If Len(Trim$(Parsed.Wojewodztwo)) = 0 Then
            Parsed.Wojewodztwo = AppContext.WojewodztwoDlaKodu(Parsed.KodPocztowy)
        End If

        If Len(Trim$(Parsed.Kraj)) = 0 Then
            Parsed.Kraj = "Polska"
        End If
    End If

    OldStreet = Nz(Contact.BusinessAddressStreet)
    OldCity = Nz(Contact.BusinessAddressCity)
    OldState = Nz(Contact.BusinessAddressState)
    OldPostalCode = Nz(Contact.BusinessAddressPostalCode)
    OldCountry = Nz(Contact.BusinessAddressCountry)

    Contact.BusinessAddressStreet = vbNullString
    Contact.BusinessAddressCity = vbNullString
    Contact.BusinessAddressState = vbNullString
    Contact.BusinessAddressPostalCode = vbNullString
    Contact.BusinessAddressCountry = vbNullString

    Contact.BusinessAddressStreet = Parsed.Ulica
    Contact.BusinessAddressCity = Parsed.Miasto
    Contact.BusinessAddressState = Parsed.Wojewodztwo
    Contact.BusinessAddressPostalCode = Parsed.KodPocztowy
    Contact.BusinessAddressCountry = Parsed.Kraj

    Contact.Save

    Summary = _
        "Adres biznesowy zosta³ uporz¹dkowany." & vbCrLf & vbCrLf & _
        "By³o:" & vbCrLf & _
        "Ulica: " & OldStreet & vbCrLf & _
        "Miasto: " & OldCity & vbCrLf & _
        "Województwo: " & OldState & vbCrLf & _
        "Kod: " & OldPostalCode & vbCrLf & _
        "Kraj: " & OldCountry & vbCrLf & vbCrLf & _
        "Jest:" & vbCrLf & _
        "Ulica: " & Contact.BusinessAddressStreet & vbCrLf & _
        "Miasto: " & Contact.BusinessAddressCity & vbCrLf & _
        "Województwo: " & Contact.BusinessAddressState & vbCrLf & _
        "Kod: " & Contact.BusinessAddressPostalCode & vbCrLf & _
        "Kraj: " & Contact.BusinessAddressCountry

    MsgBox Summary, vbInformation

End Sub

'@Description("Rozpoznaje strukturê adresu biznesowego z danych kontaktu.")
Private Function RozpoznajAdresBiznesowy(ByVal Contact As Outlook.ContactItem) As TAdresRozpoznany
Attribute RozpoznajAdresBiznesowy.VB_Description = "Rozpoznaje strukturê adresu biznesowego z danych kontaktu."

    Dim Result As TAdresRozpoznany
    Dim SourceText As String

    SourceText = ZbudujTekstZAdresuBiznesowego(Contact)
    SourceText = NormalizujTekstWielowierszowy(SourceText)

    Result.Zrodlo = SourceText
    Result.KodPocztowy = WyciagnijKodPocztowy(SourceText)
    Result.Kraj = WyciagnijKraj(SourceText, Contact)
    Result.CzyPolska = CzyPolska(Result.Kraj, SourceText, Result.KodPocztowy)
    Result.Wojewodztwo = WyciagnijWojewodztwo(SourceText)
    Result.Miasto = WyciagnijMiasto(SourceText, Result.KodPocztowy, Result.Wojewodztwo, Result.Kraj)
    Result.Ulica = WyciagnijUlice(SourceText, Result.KodPocztowy, Result.Miasto, Result.Wojewodztwo, Result.Kraj)

    Result.Ulica = NormalizujUlice(Result.Ulica)
    Result.Miasto = NormalizujMiasto(Result.Miasto)
    Result.Wojewodztwo = NormalizujProstyTekst(Result.Wojewodztwo)
    Result.Kraj = NormalizujProstyTekst(Result.Kraj)

    RozpoznajAdresBiznesowy = Result

End Function

'@Description("Buduje tekst Ÿród³owy adresu biznesowego z pól kontaktu.")
Private Function ZbudujTekstZAdresuBiznesowego(ByVal Contact As Outlook.ContactItem) As String
Attribute ZbudujTekstZAdresuBiznesowego.VB_Description = "Buduje tekst Ÿród³owy adresu biznesowego z pól kontaktu."

    Dim Text As String

    If Len(Nz(Contact.BusinessAddressStreet)) > 0 Then
        Text = Text & Contact.BusinessAddressStreet & vbCrLf
    End If

    If Len(Trim$(Nz(Contact.BusinessAddressPostalCode) & " " & Nz(Contact.BusinessAddressCity))) > 0 Then
        Text = Text & Trim$(Nz(Contact.BusinessAddressPostalCode) & " " & Nz(Contact.BusinessAddressCity)) & vbCrLf
    End If

    If Len(Nz(Contact.BusinessAddressState)) > 0 Then
        Text = Text & Contact.BusinessAddressState & vbCrLf
    End If

    If Len(Nz(Contact.BusinessAddressCountry)) > 0 Then
        Text = Text & Contact.BusinessAddressCountry & vbCrLf
    End If

    If Len(Trim$(Text)) = 0 Then
        On Error Resume Next
        Text = Nz(Contact.BusinessAddress)
        On Error GoTo 0
    End If

    ZbudujTekstZAdresuBiznesowego = Trim$(Text)

End Function

'@Description("Normalizuje koñce linii i nadmiarowe spacje w tekœcie wielowierszowym.")
Private Function NormalizujTekstWielowierszowy(ByVal Text As String) As String
Attribute NormalizujTekstWielowierszowy.VB_Description = "Normalizuje koñce linii i nadmiarowe spacje w tekœcie wielowierszowym."

    Dim localText As String
    localText = Text
    localText = Replace(localText, vbCrLf, vbLf)
    localText = Replace(localText, vbCr, vbLf)

    Do While InStr(localText, "  ") > 0
        localText = Replace(localText, "  ", " ")
    Loop

    NormalizujTekstWielowierszowy = Trim$(localText)

End Function

'@Description("Zwraca kod pocztowy w formacie NN-NNN.")
Private Function WyciagnijKodPocztowy(ByVal Text As String) As String
Attribute WyciagnijKodPocztowy.VB_Description = "Zwraca kod pocztowy w formacie NN-NNN."

    Dim Re As Object

    Set Re = CreateObject("VBScript.RegExp")
    Re.Global = False
    Re.IgnoreCase = True
    Re.Pattern = "\b\d{2}-\d{3}\b"

    If Re.Test(Text) Then
        WyciagnijKodPocztowy = Re.Execute(Text)(0).Value
    End If

End Function

'@Description("Rozpoznaje kraj na podstawie tekstu Ÿród³owego i pól kontaktu.")
Private Function WyciagnijKraj(ByVal SourceText As String, ByVal Contact As Outlook.ContactItem) As String
Attribute WyciagnijKraj.VB_Description = "Rozpoznaje kraj na podstawie tekstu Ÿród³owego i pól kontaktu."

    Dim TextLower As String
    Dim CountryField As String

    CountryField = Trim$(Nz(Contact.BusinessAddressCountry))
    If Len(CountryField) > 0 Then
        WyciagnijKraj = CountryField
        Exit Function
    End If

    TextLower = LCase$(SourceText)

    If InStr(TextLower, "polska") > 0 Then
        WyciagnijKraj = "Polska"
    ElseIf InStr(TextLower, "poland") > 0 Then
        WyciagnijKraj = "Polska"
    Else
        WyciagnijKraj = vbNullString
    End If

End Function

'@Description("Okreœla, czy adres nale¿y traktowaæ jako polski.")
Private Function CzyPolska(ByVal Country As String, ByVal SourceText As String, ByVal PostalCode As String) As Boolean
Attribute CzyPolska.VB_Description = "Okreœla, czy adres nale¿y traktowaæ jako polski."

    Dim CountryLower As String

    CountryLower = LCase$(Trim$(Country))

    If CountryLower = "polska" Or CountryLower = "poland" Then
        CzyPolska = True
        Exit Function
    End If

    If Len(CountryLower) > 0 Then
        CzyPolska = False
        Exit Function
    End If

    If InStr(LCase$(SourceText), "polska") > 0 Or InStr(LCase$(SourceText), "poland") > 0 Then
        CzyPolska = True
        Exit Function
    End If

    If Len(PostalCode) > 0 Then
        CzyPolska = True
    End If

End Function

'@Description("Rozpoznaje województwo z tekstu Ÿród³owego.")
Private Function WyciagnijWojewodztwo(ByVal Text As String) As String
Attribute WyciagnijWojewodztwo.VB_Description = "Rozpoznaje województwo z tekstu Ÿród³owego."

    Dim Wojewodztwa As Variant
    Dim i As Long
    Dim TextLower As String

    Wojewodztwa = Array( _
        "dolnoœl¹skie", "kujawsko-pomorskie", "lubelskie", "lubuskie", _
        "³ódzkie", "ma³opolskie", "mazowieckie", "opolskie", _
        "podkarpackie", "podlaskie", "pomorskie", "œl¹skie", _
        "œwiêtokrzyskie", "warmiñsko-mazurskie", "wielkopolskie", "zachodniopomorskie")

    TextLower = LCase$(Text)

    For i = LBound(Wojewodztwa) To UBound(Wojewodztwa)
        If InStr(TextLower, Wojewodztwa(i)) > 0 Then
            WyciagnijWojewodztwo = CStr(Wojewodztwa(i))
            Exit Function
        End If
    Next i

End Function

'@Description("Rozpoznaje miejscowoœæ z tekstu Ÿród³owego.")
'@Ignore ParameterNotUsed
Private Function WyciagnijMiasto( _
    ByVal Text As String, _
    ByVal PostalCode As String, _
    ByVal State As String, _
    ByVal Country As String _
) As String
Attribute WyciagnijMiasto.VB_Description = "Rozpoznaje miejscowoœæ z tekstu Ÿród³owego."

    Dim Lines() As String
    Dim i As Long
    Dim LineText As String
    Dim Re As Object

    Lines = Split(Text, vbLf)

    Set Re = CreateObject("VBScript.RegExp")
    Re.Global = False
    Re.IgnoreCase = True

    Re.Pattern = "\b\d{2}-\d{3}\b\s+(.+)$"
    For i = LBound(Lines) To UBound(Lines)
        LineText = Trim$(Lines(i))
        If Re.Test(LineText) Then
            WyciagnijMiasto = UsunDodatkiZMiasta(Re.Execute(LineText)(0).SubMatches(0), State, Country)
            Exit Function
        End If
    Next i

    Re.Pattern = "^(.+?)\s+\b\d{2}-\d{3}\b$"
    For i = LBound(Lines) To UBound(Lines)
        LineText = Trim$(Lines(i))
        If Re.Test(LineText) Then
            WyciagnijMiasto = UsunDodatkiZMiasta(Re.Execute(LineText)(0).SubMatches(0), State, Country)
            Exit Function
        End If
    Next i

    For i = LBound(Lines) To UBound(Lines)
        LineText = Trim$(Lines(i))
        If Len(LineText) > 0 Then
            If InStr(LCase$(LineText), LCase$(State)) > 0 Or InStr(LCase$(LineText), LCase$(Country)) > 0 Then
                LineText = Replace(LineText, State, vbNullString, , , vbTextCompare)
                LineText = Replace(LineText, Country, vbNullString, , , vbTextCompare)
                LineText = Replace(LineText, ",", " ")
                LineText = Trim$(NormalizujProstyTekst(LineText))

                If Len(LineText) > 0 And Not CzyLiniaUlica(LineText) And InStr(LineText, "-") = 0 Then
                    WyciagnijMiasto = LineText
                    Exit Function
                End If
            End If
        End If
    Next i

End Function

'@Description("Rozpoznaje ulicê z tekstu Ÿród³owego.")
Private Function WyciagnijUlice( _
    ByVal Text As String, _
    ByVal PostalCode As String, _
    ByVal City As String, _
    ByVal State As String, _
    ByVal Country As String _
) As String
Attribute WyciagnijUlice.VB_Description = "Rozpoznaje ulicê z tekstu Ÿród³owego."

    Dim Lines() As String
    Dim i As Long
    Dim LineText As String

    Lines = Split(Text, vbLf)

    For i = LBound(Lines) To UBound(Lines)
        LineText = Trim$(Lines(i))
        If Len(LineText) > 0 Then
            If InStr(LineText, PostalCode) = 0 _
               And InStr(LCase$(LineText), LCase$(City)) = 0 _
               And InStr(LCase$(LineText), LCase$(State)) = 0 _
               And InStr(LCase$(LineText), LCase$(Country)) = 0 Then

                If CzyLiniaUlica(LineText) Then
                    WyciagnijUlice = LineText
                    Exit Function
                End If
            End If
        End If
    Next i

    For i = LBound(Lines) To UBound(Lines)
        LineText = Trim$(Lines(i))
        If Len(LineText) > 0 Then
            If InStr(LineText, PostalCode) = 0 Then
                WyciagnijUlice = LineText
                Exit Function
            End If
        End If
    Next i

End Function

'@Description("Sprawdza, czy linia wygl¹da jak linia ulicy.")
Private Function CzyLiniaUlica(ByVal Text As String) As Boolean
Attribute CzyLiniaUlica.VB_Description = "Sprawdza, czy linia wygl¹da jak linia ulicy."

    Dim Value As String

    Value = LCase$(Trim$(Text))

    CzyLiniaUlica = _
        (Left$(Value, 3) = "ul.") Or _
        (Left$(Value, 3) = "al.") Or _
        (Left$(Value, 3) = "pl.") Or _
        (Left$(Value, 3) = "os.") Or _
        (InStr(Value, " ul.") > 0) Or _
        (InStr(Value, " al.") > 0) Or _
        (InStr(Value, " pl.") > 0) Or _
        (InStr(Value, " os.") > 0)

End Function

'@Description("Normalizuje zapis ulicy i odwraca b³êdny szyk typu '102c ul. Oœwiêcimska'.")
Private Function NormalizujUlice(ByVal Value As String) As String
Attribute NormalizujUlice.VB_Description = "Normalizuje zapis ulicy i odwraca b³êdny szyk typu '102c ul. Oœwiêcimska'."

    Dim localValue As String
    localValue = Value
    Dim Re As Object

    localValue = NormalizujProstyTekst(localValue)

    Set Re = CreateObject("VBScript.RegExp")
    Re.Global = False
    Re.IgnoreCase = True
    Re.Pattern = "^\s*([0-9]+[A-Za-z]?)\s+(ul\.|al\.|pl\.|os\.)\s+(.+?)\s*$"

    If Re.Test(localValue) Then
        localValue = Re.Replace(localValue, "$2 $3 $1")
    End If

    NormalizujUlice = Trim$(localValue)

End Function

'@Description("Normalizuje nazwê miejscowoœci.")
Private Function NormalizujMiasto(ByVal Value As String) As String
Attribute NormalizujMiasto.VB_Description = "Normalizuje nazwê miejscowoœci."

    Dim localValue As String
    localValue = Value
    localValue = NormalizujProstyTekst(localValue)
    localValue = Replace(localValue, ",", vbNullString)

    NormalizujMiasto = Trim$(localValue)

End Function

'@Description("Usuwa z nazwy miejscowoœci dodatki typu województwo i kraj.")
Private Function UsunDodatkiZMiasta(ByVal Value As String, ByVal State As String, ByVal Country As String) As String
Attribute UsunDodatkiZMiasta.VB_Description = "Usuwa z nazwy miejscowoœci dodatki typu województwo i kraj."

    Dim localValue As String
    localValue = Value
    localValue = Replace(localValue, State, vbNullString, , , vbTextCompare)
    localValue = Replace(localValue, Country, vbNullString, , , vbTextCompare)
    localValue = Replace(localValue, ",", " ")
    localValue = NormalizujProstyTekst(localValue)

    UsunDodatkiZMiasta = Trim$(localValue)

End Function

'@Description("Normalizuje zwyk³y tekst: spacje i przyciêcie.")
Private Function NormalizujProstyTekst(ByVal Value As String) As String
Attribute NormalizujProstyTekst.VB_Description = "Normalizuje zwyk³y tekst: spacje i przyciêcie."

    Dim localValue As String
    localValue = Value
    Do While InStr(localValue, "  ") > 0
        localValue = Replace(localValue, "  ", " ")
    Loop

    NormalizujProstyTekst = Trim$(localValue)

End Function

'@Description("Zwraca pusty tekst zamiast Null.")
Private Function Nz(ByVal Value As Variant) As String
Attribute Nz.VB_Description = "Zwraca pusty tekst zamiast Null."

    If IsNull(Value) Then
        Nz = vbNullString
    Else
        Nz = CStr(Value)
    End If

End Function
