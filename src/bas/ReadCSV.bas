Attribute VB_Name = "ReadCSV"
'@Folder("Kody")
Option Explicit

Public Function ReadAllText( _
    ByVal FilePath As String, _
    Optional ByVal Charset As String = "utf-8" _
) As String
    
    Dim localFilePath As String
    localFilePath = FilePath
    Dim Stream As Object
    
    localFilePath = Trim$(localFilePath)
    
    If Len(localFilePath) = 0 Then
        Err.Raise vbObjectError + 3000, "ReadAllText", "Pusta œcie¿ka pliku."
    End If
    
    If Dir$(localFilePath) = vbNullString Then
        Err.Raise vbObjectError + 3001, "ReadAllText", "Nie znaleziono pliku: " & localFilePath
    End If
    
    Set Stream = CreateObject("ADODB.Stream")
    
    On Error GoTo CleanFail
    
    With Stream
        .Type = 1              ' adTypeBinary
        .Open
        .LoadFromFile localFilePath
        .Position = 0
        .Type = 2              ' adTypeText
        .Charset = Charset
        ReadAllText = .ReadText(-1)
        .Close
    End With
    
    Set Stream = Nothing
    Exit Function

CleanFail:
    On Error Resume Next
    If Not Stream Is Nothing Then
        If Stream.State <> 0 Then Stream.Close
    End If
    Set Stream = Nothing
    
    Err.Raise vbObjectError + 3002, "ReadAllText", _
              "Nie uda³o siê odczytaæ pliku """ & localFilePath & """ jako " & Charset & "." & vbCrLf & _
              "Szczegó³y: " & Err.Description
End Function

'@Description("Dzieli tekst na linie niezale¿nie od stylu zakoñczeñ wiersza.")
Public Function SplitLines(ByVal Text As String) As String()
Attribute SplitLines.VB_Description = "Dzieli tekst na linie niezale¿nie od stylu zakoñczeñ wiersza."
    
    Dim Normalized As String
    
    Normalized = Text
    Normalized = Replace(Normalized, vbCrLf, vbLf)
    Normalized = Replace(Normalized, vbCr, vbLf)
    
    SplitLines = Split(Normalized, vbLf)
    
End Function

'@Description("Sprawdza poprawnoœæ œcie¿ki do pliku.")
'@EntryPoint
Public Sub ValidateFilePath(ByVal Value As String, ByVal ParamName As String)
Attribute ValidateFilePath.VB_Description = "Sprawdza poprawnoœæ œcie¿ki do pliku."
    
    If Len(Trim$(Value)) = 0 Then
        Err.Raise vbObjectError + 3000, "M_PlikTekstowy.ValidateFilePath", _
                  ParamName & " nie mo¿e byæ pusty."
    End If
    
    If Dir$(Value) = vbNullString Then
        Err.Raise vbObjectError + 3001, "M_PlikTekstowy.ValidateFilePath", _
                  "Nie znaleziono pliku: " & Value
    End If
    
End Sub
