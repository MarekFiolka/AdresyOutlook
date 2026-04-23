Attribute VB_Name = "NaprawiaczAdresowZaznaczenia"
'@Folder("Adresy")
Option Explicit

'@Description("Naprawia adresy biznesowe dla zaznaczonych kontaktow w biezacym folderze Outlook.")
'@EntryPoint
Public Sub NaprawAdresyBiznesoweZaznaczonychKontaktow()
Attribute NaprawAdresyBiznesoweZaznaczonychKontaktow.VB_Description = "Naprawia adresy biznesowe dla zaznaczonych kontaktow w biezacym folderze Outlook."

    Dim Explorer As Outlook.Explorer
    Dim Selection As Outlook.Selection
    Dim Item As Object
    Dim Contact As Outlook.ContactItem
    Dim Report As String

    Dim ProcessedCount As Long
    Dim SuccessCount As Long
    Dim FailedCount As Long
    Dim SkippedCount As Long
    Dim UnchangedCount As Long
    Dim WasChanged As Boolean

    Dim Summary As String
    Dim FailedDetails As String

    
    On Error Resume Next
    Set Explorer = Application.ActiveExplorer
    On Error GoTo 0

    If Explorer Is Nothing Then
        MsgBox "Brak aktywnego okna eksploratora Outlook.", vbExclamation
        Exit Sub
    End If

    Set Selection = Explorer.Selection

    If Selection Is Nothing Then
        MsgBox "Nie udalo sie pobrac zaznaczenia.", vbExclamation
        Exit Sub
    End If

    If Selection.Count = 0 Then
        MsgBox "Nie zaznaczono zadnych elementow.", vbExclamation
        Exit Sub
    End If

    For Each Item In Selection
        If TypeName(Item) = "ContactItem" Then
            Set Contact = Item
            ProcessedCount = ProcessedCount + 1

            Report = vbNullString

            WasChanged = False
            Report = vbNullString
            
            If NaprawAdresBiznesowy(Contact, True, Report, WasChanged) Then
                If WasChanged Then
                    SuccessCount = SuccessCount + 1
                Else
                    UnchangedCount = UnchangedCount + 1
                End If
            Else
                FailedCount = FailedCount + 1
                FailedDetails = FailedDetails & _
                    "- " & NazwaKontaktu(Contact) & ": " & Report & vbCrLf
            End If
        Else
            SkippedCount = SkippedCount + 1
        End If
    Next Item

    Summary = _
        "Zakonczono seryjna korekte adresow." & vbCrLf & vbCrLf & _
        "Przetworzone kontakty: " & ProcessedCount & vbCrLf & _
        "Poprawione: " & SuccessCount & vbCrLf & _
        "Bez zmian: " & UnchangedCount & vbCrLf & _
        "Bledy: " & FailedCount & vbCrLf & _
        "Pominiete elementy niebedace kontaktami: " & SkippedCount
        
    If Len(FailedDetails) > 0 Then
        Summary = Summary & vbCrLf & vbCrLf & "Szczegoly bledow:" & vbCrLf & FailedDetails
    End If

    MsgBox Summary, IIf(FailedCount = 0, vbInformation, vbExclamation)

End Sub

'@Description("Zwraca czytelna nazwe kontaktu do raportu.")
Private Function NazwaKontaktu(ByVal Contact As Outlook.ContactItem) As String
Attribute NazwaKontaktu.VB_Description = "Zwraca czytelna nazwe kontaktu do raportu."

    Dim Result As String

    Result = Trim$(Nz(Contact.FullName))
    If Len(Result) = 0 Then Result = Trim$(Nz(Contact.CompanyName))
    If Len(Result) = 0 Then Result = Trim$(Nz(Contact.FileAs))
    If Len(Result) = 0 Then Result = "(bez nazwy)"

    NazwaKontaktu = Result

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
