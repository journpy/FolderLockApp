Imports System.IO
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class Form1
    ' Declarations
    Dim fbd As FolderBrowserDialog = New FolderBrowserDialog()
    Public status As String
    Private arr As String() = New String(5) {}
    Private Sub Folder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next
        status = ""

        arr(0) = ".{2559a1f2-21d7-11d4-bdaf-00c04f60b9f0}"

        arr(1) = ".{21EC2020-3AEA-1069-A2DD-08002B30309D}"

        arr(2) = ".{2559a1f4-21d7-11d4-bdaf-00c04f60b9f0}"

        arr(3) = ".{645FF040-5081-101B-9F08-00AA002F954E}"

        arr(4) = ".{2559a1f1-21d7-11d4-bdaf-00c04f60b9f0}"

        arr(5) = ".{7007ACC7-3202-11D1-AAD2-00805FC1270E}"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Lock Folder conditions
        status = arr(0)

        If fbd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            Dim di As DirectoryInfo = New DirectoryInfo(fbd.SelectedPath)

            Dim selectedpath As String = di.Parent.FullName + di.Name

            If fbd.SelectedPath.LastIndexOf(".{") = -1 Then

                If (Not di.Root.Equals(di.Parent.FullName)) Then

                    di.MoveTo(di.Parent.FullName & "\" & di.Name & status)

                Else

                    di.MoveTo(di.Parent.FullName + di.Name & status)

                End If

                TextBox1.Text = fbd.SelectedPath
                TextBox2.Text = di.Name
                ' Call Refresh Desktop class
                Win32Helper.NotifyFileAssociationChanged()
               

            Else

                status = getstatus(status)


                di.MoveTo(fbd.SelectedPath.Substring(0, fbd.SelectedPath.LastIndexOf(".")))

                TextBox1.Text = fbd.SelectedPath.Substring(0, fbd.SelectedPath.LastIndexOf("."))
                TextBox2.Text = di.Name
                ' Call Refresh Desktop class
                Win32Helper.NotifyFileAssociationChanged()
            End If

        End If


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Unlock Folder conditions
        status = arr(1)

        If fbd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            Dim di As DirectoryInfo = New DirectoryInfo(fbd.SelectedPath)

            Dim selectedpath As String = di.Parent.FullName + di.Name

            If fbd.SelectedPath.LastIndexOf(".{") = -1 Then

                If (Not di.Root.Equals(di.Parent.FullName)) Then

                    di.MoveTo(di.Parent.FullName & "\" & di.Name & status)

                Else

                    di.MoveTo(di.Parent.FullName + di.Name & status)

                End If

                TextBox1.Text = fbd.SelectedPath
                TextBox2.Text = di.Name
                ' Call Refresh Desktop class
                Win32Helper.NotifyFileAssociationChanged()


            Else

                status = getstatus(status)

                di.MoveTo(fbd.SelectedPath.Substring(0, fbd.SelectedPath.LastIndexOf(".")))

                TextBox1.Text = fbd.SelectedPath.Substring(0, fbd.SelectedPath.LastIndexOf("."))
                TextBox2.Text = di.Name
                ' Call Refresh Desktop class
                Win32Helper.NotifyFileAssociationChanged()
            End If

        End If

    End Sub
    Private Function getstatus(ByVal stat As String) As String

        For i As Integer = 0 To 5

            If stat.LastIndexOf(arr(i)) <> -1 Then

                stat = stat.Substring(stat.LastIndexOf("."))

            End If

        Next i

        Return stat

    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Colouring the form
        Me.BackColor = Color.Silver
        Me.ForeColor = Color.Blue
    End Sub
End Class
