Imports System.IO

Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread


    Public filefolder As String = ""
    Public textcase As String = "lowercase"
    Public forceextension As Boolean = True
    Public forcefoldername As Boolean = False

    Private filecount As Long = 0
    Private filecounter As Long = 0

    Public result As String = ""

    Public Event WorkerComplete(ByVal Result As String)
    Public Event WorkerProgress(ByVal filesrenamed As Long, ByVal currentfilename As String, ByVal totalfiles As Long)



#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.ToString)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), ) & " - " & identifier_msg & ":" & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Mass Filename Case Changer's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Activity_Logger(ByVal identifier_msg As String)
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), ) & " - " & identifier_msg)
            filewriter.Flush()
            filewriter.Close()
        Catch ex As Exception
            Error_Handler(ex, "Activity Logger")
        End Try
    End Sub

    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    ' Sets the thread using the AddressOf the subroutine where
                    ' the thread will start.
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerExecute)
                    ' Starts the thread.
                    WorkerThread.Start()

            End Select
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerExecute()
        Try
            filecount = 0
            filecounter = 0
            RaiseEvent WorkerProgress(filecounter, "", filecount)
            ' filecount = filecount + Directory.GetFiles(filefolder).Length


            FolderWalker(filefolder)


            '        RaiseEvent WorkerProgress(runner, oldname & " -> " & finfo.Name, dinfo.GetFiles.Length)
            result = "Success" & Directory.GetFiles(filefolder).Length
            RaiseEvent WorkerComplete(result)
        Catch ex As Exception
            result = "Failure"
            RaiseEvent WorkerComplete(result)
        End Try

        WorkerThread.Abort()
    End Sub

    Private Sub FolderWalker(ByVal targetDirectory As String)
        Try
            Dim dinfo As DirectoryInfo = New DirectoryInfo(targetDirectory)
            Dim fileEntries As String() = Directory.GetFiles(targetDirectory)
            filecount = filecount + Directory.GetFiles(targetDirectory).Length
            Dim fileName As String
            For Each fileName In fileEntries
                Try
                    Dim finfo As FileInfo = New FileInfo(fileName)
                    Dim oldname, oldext As String
                    oldname = finfo.Name.Substring(0, finfo.Name.Length - finfo.Extension.Length)
                    Dim compname As String = oldname
                    oldext = finfo.Extension
                    Dim compext As String = oldext
                    Select Case textcase
                        Case "lowercase"
                            oldname = oldname.ToLower
                        Case "uppercase"
                            oldname = oldname.ToUpper
                        Case "titlecase"
                            Dim str As String() = oldname.Split(" ")
                            Dim newname As String = ""
                            For Each stri As String In str
                                If stri.Length > 1 Then
                                    newname = newname & stri.Substring(0, 1).ToUpper & stri.Substring(1, stri.Length - 1).ToLower & " "
                                Else
                                    newname = newname & stri.Substring(0, 1).ToUpper & " "
                                End If
                            Next
                            newname = newname.Substring(0, newname.Length - 1)
                            oldname = newname
                    End Select
                    If forceextension = True Then
                        oldext = oldext.ToLower
                    End If

                    If Not compname = oldname Or Not oldext = compext Then
                        filecounter = filecounter + 1
                        RaiseEvent WorkerProgress(filecounter, finfo.FullName, filecount)
                        Activity_Logger("RENAMING FILE NECESSARY")
                        Activity_Logger("Old Name: " & finfo.FullName)
                        Activity_Logger("New Name: " & (finfo.DirectoryName & "\").Replace("\\", "\") & oldname & ("." & oldext).Replace("..", "."))
                        finfo.MoveTo((finfo.DirectoryName & "\").Replace("\\", "\") & oldname & ("." & oldext).Replace("..", "."))

                    End If

                    finfo = Nothing

                Catch ex As Exception
                    Error_Handler(ex, "Processing Files")
                End Try

            Next fileName

            Dim subdirectoryEntries As String() = Directory.GetDirectories(targetDirectory)
            Dim subdirectory As String
            For Each subdirectory In subdirectoryEntries

                FolderWalker(subdirectory)
            Next subdirectory


            If forcefoldername = True Then


            Dim oldname, oldext As String
            oldname = dinfo.Name
            Dim compname As String = oldname

            Select Case textcase
                Case "lowercase"
                    oldname = oldname.ToLower
                Case "uppercase"
                    oldname = oldname.ToUpper
                Case "titlecase"
                    Dim str As String() = oldname.Split(" ")
                    Dim newname As String = ""
                    For Each stri As String In str
                        If stri.Length > 1 Then
                            newname = newname & stri.Substring(0, 1).ToUpper & stri.Substring(1, stri.Length - 1).ToLower & " "
                        Else
                            newname = newname & stri.Substring(0, 1).ToUpper & " "
                        End If
                    Next
                    newname = newname.Substring(0, newname.Length - 1)
                    oldname = newname
            End Select

            If Not compname = oldname Then
                    'filecounter = filecounter + 1
                    'RaiseEvent WorkerProgress(filecounter, dinfo.FullName, filecount)
                Activity_Logger("RENAMING FOLDER NECESSARY")
                    Activity_Logger("Old Name: " & dinfo.FullName)
                    dinfo.MoveTo((dinfo.Parent.FullName & "\").Replace("\\", "\") & oldname & "_X1X1X_")
                    dinfo.MoveTo((dinfo.Parent.FullName & "\").Replace("\\", "\") & oldname)
                Activity_Logger("New Name: " & dinfo.FullName)
            End If
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub 'ProcessDirectory




End Class
