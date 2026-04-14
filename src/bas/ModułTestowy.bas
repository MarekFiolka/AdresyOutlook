Attribute VB_Name = "Modu³Testowy"
'@Folder("Test")
Option Explicit

'@EntryPoint
Public Sub TestAppContext()

    Debug.Print AppContext.ToDebugString
    Debug.Print AppContext.WojewodztwoDlaKodu("00-001")
    Debug.Print AppContext.MiejscowoscDlaKodu("00-001")
    Debug.Print AppContext.KodyPocztowe.Count

End Sub
