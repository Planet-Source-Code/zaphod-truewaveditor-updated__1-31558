Attribute VB_Name = "Module1"
'WavEditor
'by: Paul Bryan in 2002
'Allows for graphical isolation of sample ranges
'will Paste Selected Data Into New Files

' I hope this helps, feel free to re-use this code.

Public SCount As Integer ' Multiple Open Wave Files
Public Scope(255) As Form ' Session Filecount
Public Sub main()
    MDIMain.Show
    frmAbout.StartFlash
End Sub
Public Sub LoadNewFile(FName As String) ' Open another file
        SCount = SCount + 1
        
        Set Scope(SCount) = New WavForm
        Scope(SCount).SetFocus
        Call Scope(SCount).LoadFileData(SCount, FName)
    
        Exit Sub
End Sub

