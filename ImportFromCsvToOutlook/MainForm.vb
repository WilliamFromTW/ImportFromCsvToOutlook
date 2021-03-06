﻿Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook

Public Class MainForm
    Dim ffile As OpenFileDialog
    Dim outlook = GetApplicationObject()
    Dim ns As Outlook.NameSpace = outlook.GetNamespace("MAPI")
    Dim aFolder = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts)
    Dim afolderItems = aFolder.Items

    Private Sub CreateDistributionList(strMyToken() As String)


        Dim bList As Boolean
        bList = False

        Dim iCount = afolderItems.Count
        Dim sTypeName As String
        Dim sItemObject As Outlook.ContactItem
        Dim sTemp As String

        For x = 1 To iCount
            sTypeName = TypeName(afolderItems.Item(x))



            If sTypeName = "ContactItem" Then
                sItemObject = afolderItems.Item(x)


                'Check if the distribution list exists
                If String.IsNullOrEmpty(sItemObject.FirstName) Or String.IsNullOrEmpty(strMyToken(3)) Then
                    Continue For
                End If
                sTemp = strMyToken(3).Replace("""", "")
                If strMyToken(3).Replace("""", "").Equals(sItemObject.Email1Address) Then
                    bList = True
                    Exit For
                End If
            End If
        Next x
        If Not bList Then

            Dim contact As Outlook.ContactItem = afolderItems.Add(OlItemType.olContactItem)
            Dim strDisplayName, strFirstName, strLastName, strEmail, strTitle, strDepartment, strOfficeLocation As String

            If strMyToken(0) IsNot Nothing Then
                strDisplayName = strMyToken(0).Replace("""", "").Trim()
            Else
                strDisplayName = ""
            End If
            If strMyToken(1) IsNot Nothing And Not strMyToken(1).Equals("") Then
                strFirstName = strMyToken(1).Replace("""", "").Trim()
            Else
                strFirstName = strDisplayName
            End If
            If strMyToken(2) IsNot Nothing And Not strMyToken(2).Equals("") Then
                strLastName = strMyToken(2).Replace("""", "").Trim()
            Else
                strLastName = ""
            End If
            If strMyToken(4) IsNot Nothing And Not strMyToken(3).Equals("") Then
                strEmail = strMyToken(3).Replace("""", "").Trim()
            Else
                strEmail = ""
            End If
            If strMyToken(4) IsNot Nothing And Not strMyToken(4).Equals("") Then
                strTitle = strMyToken(4).Replace("""", "").Trim()
            Else
                strTitle = ""
            End If
            If strMyToken(5) IsNot Nothing And Not strMyToken(5).Equals("") Then
                strDepartment = strMyToken(5).Replace("""", "").Trim()
            Else
                strDepartment = ""
            End If
            If strMyToken(6) IsNot Nothing And Not strMyToken(6).Equals("") Then
                strOfficeLocation = strMyToken(6).Replace("""", "").Trim()
            Else
                strOfficeLocation = ""
            End If


            With contact
                .FirstName = strFirstName
                .LastName = strLastName
                .Email1Address = strEmail
                .Title = strTitle
                .Department = strDepartment
                .OfficeLocation = strOfficeLocation
                .Save()
            End With

            Marshal.ReleaseComObject(contact)
        End If

    End Sub
    Function GetApplicationObject() As Outlook.Application
        Dim application As Outlook.Application
        application = New Outlook.Application()
        Return application
    End Function
    Function getFileContentsFromHTTP(csvfile As String) As StringReader

        ' bypass private sign cert files
        ServicePointManager.ServerCertificateValidationCallback = Function(s, c, h, e) True
        Dim strReader As New StringReader("")
        Using client As HttpClient = New HttpClient()
            Using response As HttpResponseMessage = client.GetAsync(csvfile).Result

                Using content As HttpContent = response.Content
                    ' Get contents of page as a String.
                    Dim result As String = content.ReadAsStringAsync().Result
                    If result IsNot Nothing Then
                        strReader = New StringReader(result)

                    End If
                End Using
            End Using
        End Using
        Return strReader
    End Function
    Private Sub Form1_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim iArgsLength = Environment.GetCommandLineArgs.Length
        Dim filePathstring As String = ""
        Dim fileReader As System.IO.StreamReader
        Dim strFileContents As StringReader = New StringReader("")


        Try
            If iArgsLength > 1 Then
                'MsgBox(Environment.GetCommandLineArgs(1))
                filePathstring = Environment.GetCommandLineArgs(1)
                If filePathstring.ToLower.IndexOf("http") = -1 Then
                    fileReader = My.Computer.FileSystem.OpenTextFileReader(filePathstring)
                    strFileContents = New StringReader(fileReader.ReadToEnd)
                Else
                    strFileContents = getFileContentsFromHTTP(filePathstring)
                End If
            Else
                MsgBox("import contacts to outlook (2013 or above) , please select source file(csv)")
                ffile = New OpenFileDialog
                If (ffile.ShowDialog().Equals(DialogResult.OK)) Then
                    filePathstring = ffile.FileName
                    fileReader = My.Computer.FileSystem.OpenTextFileReader(filePathstring)
                    strFileContents = New StringReader(fileReader.ReadToEnd)
                Else
                    Me.Close()
                    Return
                End If
            End If
        Catch fileException As System.Exception
            If iArgsLength = 1 Then
                MsgBox("read file failed")
            End If
            Me.Close()
            Return
        End Try

        If String.IsNullOrEmpty(filePathstring) Then
            Me.Close()
            Return
        End If

        Dim stringReader As String
        Dim iCount As Integer = 0
        While strFileContents IsNot Nothing And strFileContents.Peek <> -1

            stringReader = strFileContents.ReadLine()
            iCount += 1

            If (iCount = 1) Then
                Continue While
            End If

            'MsgBox("The first line of the file is " & stringReader)
            Dim myToken() As String = stringReader.Split(",")

            If myToken.Length <> 7 Then
                If iArgsLength = 1 Then
                    MsgBox("line # = " + CStr(iCount) + " in excel is not correct, app will close")
                End If
                Me.Close()
            Else
                CreateDistributionList(myToken)
            End If

        End While
        If iArgsLength = 1 Then
            MsgBox("Import contacts finished!")
        End If
        Me.Close()
        Return
    End Sub
End Class
