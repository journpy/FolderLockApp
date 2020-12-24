Imports System.Runtime.InteropServices

Public NotInheritable Class Win32Helper
    ' Refresh Desktop

    Const SHCNE_ASSOCCHANGED As Integer = &H8000000
    Const SHCNF_IDLIST As Integer = 0

    Private Class NativeMethods
        <DllImport("shell32")> _
        Public Shared Sub SHChangeNotify(ByVal wEventId As Integer, ByVal flags As Integer, ByVal item1 As IntPtr, ByVal item2 As IntPtr)
        End Sub
    End Class

    Public Shared Sub NotifyFileAssociationChanged()
        ' SHChangeNotify notifies the system of events. 
        ' You can notify that various events occured, one is SHCNE_ASSOCCHANGED: 
        ' "A file type association has changed.  
        ' SHCNF_IDLIST must be specified in the uFlags parameter.  
        ' dwItem1 and dwItem2 are not used and must be NULL."  
        NativeMethods.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, Nothing, Nothing)
    End Sub
End Class
