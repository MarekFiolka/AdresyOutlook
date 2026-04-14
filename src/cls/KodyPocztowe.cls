VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KodyPocztowe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule EmptyIfBlock, AnnotationInIncompatibleComponentType
'@Folder("Kody")
Option Explicit

Private Type TKodyPocztowe
    Items As Collection
    IndexByKod As Object ' Scripting.Dictionary
End Type

Private this As TKodyPocztowe

Private Sub Class_Initialize()
    Set this.Items = New Collection
    Set this.IndexByKod = CreateObject("Scripting.Dictionary")
    this.IndexByKod.CompareMode = 1 ' TextCompare
End Sub

Private Sub Class_Terminate()
    Set this.IndexByKod = Nothing
    Set this.Items = Nothing
End Sub

'@Description("Enumerator kolekcji kodów pocztowych")
'@Enumerator
'@EntryPoint
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Enumerator kolekcji kodów pocztowych"
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = this.Items.[_NewEnum]
End Property

'@Description("Licznik elementów kolekcji")
'@EntryPoint
Public Property Get Count() As Long
Attribute Count.VB_Description = "Licznik elementów kolekcji"
    Count = this.Items.Count
End Property

'@Description("Wybrany element kolekcji")
'@DefaultMember
'@EntryPoint
Public Property Get Item(ByVal Index As Variant) As KodPocztowy
Attribute Item.VB_Description = "Wybrany element kolekcji"
Attribute Item.VB_UserMemId = 0
    ValidateKey Index, "Index"
    Set Item = this.Items.Item(Index)
End Property

'@Description("Dodaje rekord kodu pocztowego do kolekcji")
Public Sub Add(ByVal Value As KodPocztowy)
Attribute Add.VB_Description = "Dodaje rekord kodu pocztowego do kolekcji"
    
    Dim Key As String
    
    ValidateObject Value, "Value"
    
    Key = NormalizeKod(Value.Kod)
    
    If Len(Key) = 0 Then
        Err.Raise vbObjectError + 1000, TypeName(Me), "Kod pocztowy nie mo¿e byæ pusty."
    End If
    
    If this.IndexByKod.Exists(Key) Then
        Err.Raise vbObjectError + 1001, TypeName(Me), "Kod pocztowy ju¿ istnieje: " & Key
    End If
    
    this.Items.Add Value, Key
    this.IndexByKod.Add Key, Value
    
End Sub

'@Description("Usuwa wszystkie elementy kolekcji")
Public Sub Clear()
Attribute Clear.VB_Description = "Usuwa wszystkie elementy kolekcji"
    Set this.Items = New Collection
    Set this.IndexByKod = CreateObject("Scripting.Dictionary")
    this.IndexByKod.CompareMode = 1
End Sub

'@Description("Sprawdza, czy kolekcja zawiera podany kod pocztowy")
Public Function Contains(ByVal Kod As String) As Boolean
Attribute Contains.VB_Description = "Sprawdza, czy kolekcja zawiera podany kod pocztowy"
    Contains = this.IndexByKod.Exists(NormalizeKod(Kod))
End Function

'@Description("Zwraca rekord dla podanego kodu pocztowego")
'@EntryPoint
Public Property Get ByKod(ByVal Kod As String) As KodPocztowy
Attribute ByKod.VB_Description = "Zwraca rekord dla podanego kodu pocztowego"
    
    Dim Key As String
    Key = NormalizeKod(Kod)
    
    If Not this.IndexByKod.Exists(Key) Then
        Err.Raise vbObjectError + 1002, TypeName(Me), "Nie znaleziono kodu pocztowego: " & Key
    End If
    
    Set ByKod = this.IndexByKod.Item(Key)
    
End Property

'@Description("Próbuje zwróciæ rekord dla podanego kodu pocztowego")
Public Function TryGetByKod(ByVal Kod As String, ByRef Result As KodPocztowy) As Boolean
Attribute TryGetByKod.VB_Description = "Próbuje zwróciæ rekord dla podanego kodu pocztowego"
    
    Dim Key As String
    Key = NormalizeKod(Kod)
    
    If this.IndexByKod.Exists(Key) Then
        Set Result = this.IndexByKod.Item(Key)
        TryGetByKod = True
    Else
        Set Result = Nothing
        TryGetByKod = False
    End If
    
End Function

'@Description("Zwraca województwo dla podanego kodu pocztowego")
'@EntryPoint
Public Function WojewodztwoDlaKodu(ByVal Kod As String) As String
Attribute WojewodztwoDlaKodu.VB_Description = "Zwraca województwo dla podanego kodu pocztowego"
    
    Dim Found As KodPocztowy
    
    If TryGetByKod(Kod, Found) Then
        WojewodztwoDlaKodu = Found.Wojewodztwo
    End If
    
End Function

'@Description("Zwraca miejscowoœæ dla podanego kodu pocztowego")
'@EntryPoint
Public Function MiejscowoscDlaKodu(ByVal Kod As String) As String
Attribute MiejscowoscDlaKodu.VB_Description = "Zwraca miejscowoœæ dla podanego kodu pocztowego"
    
    Dim Found As KodPocztowy
    
    If TryGetByKod(Kod, Found) Then
        MiejscowoscDlaKodu = Found.Miejscowosc
    End If
    
End Function

'@Description("Wczytuje dane z pliku CSV w kodowaniu UTF-8.")
Public Sub LoadFromCsv(ByVal FilePath As String)
Attribute LoadFromCsv.VB_Description = "Wczytuje dane z pliku CSV w kodowaniu UTF-8."
    
    Dim Text As String
    Dim Lines() As String
    Dim Fields() As String
    Dim i As Long
    Dim LineText As String
    Dim Entry As KodPocztowy
    
    ValidateFilePath FilePath, "FilePath"
    
    Clear
    
    Text = ReadAllText(FilePath, "utf-8")
    Lines = SplitLines(Text)
    
    For i = LBound(Lines) To UBound(Lines)
        
        LineText = Trim$(Lines(i))
        
        If i = LBound(Lines) Then
            ' Pomijamy nag³ówek
        ElseIf Len(LineText) > 0 Then
            
            Fields = ParseDelimitedLine(LineText, ";")
            
            If UBound(Fields) >= 4 Then
                
                Set Entry = New KodPocztowy
                
                Entry.Init _
                    NormalizeCsvField(Fields(0)), _
                    NormalizeCsvField(Fields(1)), _
                    NormalizeCsvField(Fields(2)), _
                    NormalizeCsvField(Fields(3)), _
                    NormalizeCsvField(Fields(4))
                
                If Len(Entry.Kod) > 0 Then
                    If Not Contains(Entry.Kod) Then
                        Add Entry
                    End If
                End If
                
            Else
                Err.Raise vbObjectError + 3002, TypeName(Me), _
                          "Nieprawid³owa liczba pól w wierszu nr " & CStr(i + 1) & "."
            End If
        End If
    Next i
    
End Sub

Private Function NormalizeKod(ByVal Value As String) As String
    NormalizeKod = Trim$(Value)
End Function

Private Sub ValidateKey(ByVal Value As Variant, ByVal ParamName As String)
    
    If IsObject(Value) Then
        Err.Raise vbObjectError + 1010, TypeName(Me), ParamName & " nie mo¿e byæ obiektem."
    End If
    
    If IsEmpty(Value) Then
        Err.Raise vbObjectError + 1011, TypeName(Me), ParamName & " jest wymagany."
    End If
    
End Sub

Private Sub ValidateObject(ByVal Value As Object, ByVal ParamName As String)
    
    If Value Is Nothing Then
        Err.Raise vbObjectError + 1012, TypeName(Me), ParamName & " nie mo¿e byæ Nothing."
    End If
    
End Sub

Private Sub ValidateFilePath(ByVal Value As String, ByVal ParamName As String)
    
    If Len(Trim$(Value)) = 0 Then
        Err.Raise vbObjectError + 1013, TypeName(Me), ParamName & " nie mo¿e byæ pusty."
    End If
    
    If Dir$(Value) = vbNullString Then
        Err.Raise vbObjectError + 1014, TypeName(Me), "Nie znaleziono pliku: " & Value
    End If
    
End Sub
