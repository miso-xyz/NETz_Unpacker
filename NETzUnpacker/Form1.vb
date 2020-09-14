Imports System.IO
Imports System.Drawing.Text
Imports System.Resources
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports ICSharpCode.SharpZipLib.Zip.Compression.Streams
Public Class Form1
    Private Shared rm As ResourceManager
    Private Shared inResourceResolveFlag As Boolean = False
    Private Shared xrRm As ArrayList
    Private Shared targettedAssembly As String
    Private Shared myFonts As PrivateFontCollection
    Private Shared fontBuffer As IntPtr

    Public Sub New()
        InitializeComponent()

        If myFonts Is Nothing Then
            myFonts = New PrivateFontCollection()
            Dim font As Byte() = My.Resources.font
            fontBuffer = Marshal.AllocCoTaskMem(font.Length)
            Marshal.Copy(font, 0, fontBuffer, font.Length)
            myFonts.AddMemoryFont(fontBuffer, font.Length)
        End If
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim fam As FontFamily = myFonts.Families(0)
        Label1.Font = New Font(fam, 6.0F)
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.AllowDrop = True
    End Sub

    Private Sub Form1_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each path In files
            targettedAssembly = path
            Label1.Text = "unpacker by misonothx" & vbCrLf & vbCrLf & "tested on NETz 0.4.8" & vbCrLf & vbCrLf & "sinister.ly <3" & vbCrLf & vbCrLf & "---------------" & vbCrLf & vbCrLf & "loading..."
            StartApp()
        Next
    End Sub
    Public Function StartApp() As Integer
        Dim resource As Byte()
        Try
            resource = GetResource("A6C24BF5-3690-4982-887E-11E1B159B249")
        Catch ex As Exception
            Label1.Text = "unpacker by misonothx" & vbCrLf & vbCrLf & "tested on NETz 0.4.8" & vbCrLf & vbCrLf & "sinister.ly <3" & vbCrLf & vbCrLf & "---------------" & vbCrLf & vbCrLf & "resource not found"
            Return 0
        End Try
        Try
            GetAssembly(resource)
            Label1.Text = "unpacker by misonothx" & vbCrLf & vbCrLf & "tested on NETz 0.4.8" & vbCrLf & vbCrLf & "sinister.ly <3" & vbCrLf & vbCrLf & "---------------" & vbCrLf & vbCrLf & "extracting..."
        Catch ex As Exception
            Label1.Text = "unpacker by misonothx" & vbCrLf & vbCrLf & "tested on NETz 0.4.8" & vbCrLf & vbCrLf & "sinister.ly <3" & vbCrLf & vbCrLf & "---------------" & vbCrLf & vbCrLf & "failed to extract"
            Return 0
        End Try
        Label1.Text = "unpacker by misonothx" & vbCrLf & vbCrLf & "tested on NETz 0.4.8" & vbCrLf & vbCrLf & "sinister.ly <3" & vbCrLf & vbCrLf & "---------------" & vbCrLf & vbCrLf & "extracted!"
        Timer1.Start()
        Return 1
    End Function
    Private Sub Form1_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Function MangleDllName(ByVal dll As String) As String
        Dim text As String = dll.Replace(" ", "!1")
        text = text.Replace(",", "!2")
        text = text.Replace(".Resources", "!3")
        text = text.Replace(".resources", "!3")
        Return text.Replace("Culture", "!4")
    End Function

    Private Function GetResource(ByVal id As String) As Byte()
        Dim array As Byte() = Nothing
        If rm Is Nothing Then
            rm = New ResourceManager("app", Assembly.LoadFile(targettedAssembly))
        End If
        Try
            inResourceResolveFlag = True
            Dim name As String = MangleDllName(id)
            If array Is Nothing AndAlso xrRm IsNot Nothing Then
                For i As Integer = 0 To xrRm.Count - 1
                    Try
                        Dim resourceManager As ResourceManager = CType(xrRm(i), ResourceManager)
                        If resourceManager IsNot Nothing Then
                            array = CType(resourceManager.GetObject(name), Byte())
                        End If
                    Catch
                    End Try
                    If array IsNot Nothing Then
                        Exit For
                    End If
                Next
            End If
            If array Is Nothing Then
                array = CType(rm.GetObject(name), Byte())
            End If
        Finally
            inResourceResolveFlag = False
        End Try
        Return array
    End Function

    Private Function GetAssembly(ByVal data As Byte()) As Assembly
        Dim memoryStream As MemoryStream = Nothing
        Dim result As Assembly = Nothing
        Try
            memoryStream = UnZip(data)
            memoryStream.Seek(0L, SeekOrigin.Begin)
            If IO.Directory.Exists("NETz_Unpacker") = False Then
                IO.Directory.CreateDirectory("NETz_Unpacker")
            End If
            IO.File.WriteAllBytes("NETz_Unpacker/" & IO.Path.GetFileName(targettedAssembly), memoryStream.ToArray())
            result = Assembly.Load(memoryStream.ToArray())
        Finally
            If memoryStream IsNot Nothing Then
                memoryStream.Close()
            End If
            memoryStream = Nothing
        End Try
        Return result
    End Function

    Private Function UnZip(ByVal data As Byte()) As MemoryStream
        If data Is Nothing Then
            Return Nothing
        End If
        Dim memoryStream As MemoryStream = Nothing
        Dim memoryStream2 As MemoryStream = Nothing
        Dim inflaterInputStream As InflaterInputStream = Nothing
        Try
            memoryStream = New MemoryStream(data)
            memoryStream2 = New MemoryStream()
            inflaterInputStream = New InflaterInputStream(memoryStream)
            Dim array As Byte() = New Byte(data.Length - 1) {}
            While True
                Dim num As Integer = inflaterInputStream.Read(array, 0, array.Length)
                If num <= 0 Then
                    Exit While
                End If
                memoryStream2.Write(array, 0, num)
            End While
            memoryStream2.Flush()
            memoryStream2.Seek(0L, SeekOrigin.Begin)
        Finally
            If memoryStream IsNot Nothing Then
                memoryStream.Close()
            End If
            If inflaterInputStream IsNot Nothing Then
                inflaterInputStream.Close()
            End If
            memoryStream = Nothing
            inflaterInputStream = Nothing
        End Try
        Return memoryStream2
    End Function

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label1.Text = "unpacker by misonothx" & vbCrLf & vbCrLf & "tested on NETz 0.4.8" & vbCrLf & vbCrLf & "sinister.ly <3" & vbCrLf & vbCrLf & "---------------" & vbCrLf & vbCrLf & "drop file"
        Timer1.Stop()
    End Sub
End Class

