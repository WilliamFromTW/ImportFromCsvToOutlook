Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook

Public Class MainForm
    Dim ffile As OpenFileDialog

    Private Function CreateDistributionList(strDisplayName As String, strEmail As String)

        Dim outlook = GetApplicationObject()
        Dim ns As Outlook.NameSpace = outlook.GetNamespace("MAPI")
        Dim aFolder = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts)
        Dim afolderItems = aFolder.Items
        Dim bList As Boolean
        bList = False
        Dim bMember As Boolean
        bMember = False

        Dim iCount = afolderItems.Count
        Dim olDistList As Object
        For x = 1 To iCount
            If TypeName(afolderItems.Item(x)) = "DistListItem" Then
                olDistList = afolderItems.Item(x)
                'Check if the distribution list exists
                If olDistList.DLName = strDisplayName Then
                    bList = True
                End If
            End If
        Next x
        If Not bList Then

            Dim contact As Outlook.ContactItem = afolderItems.Add(OlItemType.olContactItem)

            With contact
                .FirstName = strDisplayName
                .Email1Address = strEmail
                .Save()
            End With

            Marshal.ReleaseComObject(contact)
        End If
    End Function
    Function GetApplicationObject() As Outlook.Application
        Dim application As Outlook.Application
        application = New Outlook.Application()
        Return application
    End Function

    Private Sub Form1_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load

        ffile = New OpenFileDialog
        If (ffile.ShowDialog().Equals(DialogResult.OK)) Then
            Dim filePathstring = ffile.FileName

            Dim fileReader As System.IO.StreamReader
            fileReader = My.Computer.FileSystem.OpenTextFileReader(filePathstring)
            Dim stringReader As String
            Dim iCount As Integer = 0
            While Not fileReader.EndOfStream
                stringReader = fileReader.ReadLine()
                iCount += 1

                If (iCount = 1) Then
                    Continue While
                End If

                'MsgBox("The first line of the file is " & stringReader)
                Dim myToken() As String = stringReader.Split(",")

                If myToken.Length <> 2 Then
                    MsgBox("line # = " + CStr(iCount) + " in excel is not correct, app will close")
                    Me.Close()
                Else
                    CreateDistributionList(myToken(0), myToken(1))
                End If

            End While
            Me.Close()

        Else
            Me.Close()


        End If

    End Sub
End Class
