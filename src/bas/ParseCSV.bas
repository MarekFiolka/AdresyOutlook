Attribute VB_Name = "ParseCSV"
'@Folder("Kody")
Option Explicit

'@Description("Parsuje pojedynczy wiersz pliku rozdzielanego wskazanym separatorem z obsluga cudzyslowow.")
Public Function ParseDelimitedLine( _
    ByVal LineText As String, _
    Optional ByVal Delimiter As String = ";" _
) As String()
Attribute ParseDelimitedLine.VB_Description = "Parsuje pojedynczy wiersz pliku rozdzielanego wskazanym separatorem z obs³ug¹ cudzys³owów."

    Dim Result() As String
    Dim Buffer As String
    Dim i As Long
    Dim Ch As String
    Dim InQuotes As Boolean
    Dim FieldIndex As Long
    
    If Len(Delimiter) <> 1 Then
        Err.Raise vbObjectError + 2000, "M_Csv.ParseDelimitedLine", _
                  "Delimiter musi byc pojedynczym znakiem."
    End If
    
    ReDim Result(0 To 0)
    Buffer = vbNullString
    FieldIndex = 0
    InQuotes = False
    
    For i = 1 To Len(LineText)
        
        Ch = Mid$(LineText, i, 1)
        
        Select Case Ch
            
            Case """"
                If InQuotes Then
                    If i < Len(LineText) And Mid$(LineText, i + 1, 1) = """" Then
                        Buffer = Buffer & """"
                        i = i + 1
                    Else
                        InQuotes = False
                    End If
                Else
                    InQuotes = True
                End If
            
            Case Delimiter
                If InQuotes Then
                    Buffer = Buffer & Ch
                Else
                    Result(FieldIndex) = Buffer
                    Buffer = vbNullString
                    FieldIndex = FieldIndex + 1
                    ReDim Preserve Result(0 To FieldIndex)
                End If
            
            Case Else
                Buffer = Buffer & Ch
        
        End Select
    Next i
    
    Result(FieldIndex) = Buffer
    
    ParseDelimitedLine = Result

End Function

'@Description("Usuwa zewnetrzne cudzyslowy z pola i normalizuje zapis podwojnych cudzyslowow.")
Public Function NormalizeCsvField(ByVal Value As String) As String
Attribute NormalizeCsvField.VB_Description = "Usuwa zewnêtrzne cudzys³owy z pola i normalizuje zapis podwójnych cudzys³owów."
    
    Dim Text As String
    Text = Trim$(Value)
    
    If Len(Text) >= 2 Then
        If Left$(Text, 1) = """" And Right$(Text, 1) = """" Then
            Text = Mid$(Text, 2, Len(Text) - 2)
        End If
    End If
    
    '@Ignore AssignmentNotUsed
    Text = Replace(Text, """""", """")
    
    NormalizeCsvField = Trim$(Text)

End Function
