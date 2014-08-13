Imports System.Drawing.Printing
Imports System.IO


Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Private splash_loader As Splash_Screen
    Public dataloaded As Boolean = False

    Private printFont As Font
    Private streamToPrint As StreamReader
    Private Shared filePath As String


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call


    End Sub

    Public Sub New(ByVal splash As Splash_Screen)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        splash_loader = splash

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog





    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.Label4 = New System.Windows.Forms.Label
        Me.Button6 = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.SuspendLayout()
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.LemonChiffon
        Me.Label4.Location = New System.Drawing.Point(384, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "BUILD 20051007.1"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.Black
        Me.Button6.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button6.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.ForeColor = System.Drawing.Color.Yellow
        Me.Button6.Location = New System.Drawing.Point(352, 24)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(88, 24)
        Me.Button6.TabIndex = 10
        Me.Button6.Text = "GET FILE LIST"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(8, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(280, 48)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "SINGLE LEVEL FILE LISTER IS AN APPLICATION TO QUICKLY GIVE YOU A TEXT FILE AND PR" & _
        "INTOUT OF ALL FILE NAMES CONTAINED WITHIN A SPECIFIED FOLDER. YOU CAN CHOOSE WHE" & _
        "THER OR NOT TO DISPLAY FULLY QUALIFIED PATHS OR SIMPLE FILE NAMES."
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "txt"
        Me.SaveFileDialog1.FileName = "filelist.txt"
        Me.SaveFileDialog1.Filter = "Text files|*.txt|All files|*.*"
        Me.SaveFileDialog1.Title = "Select the savepath for the generated for the file listing"
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.Description = "Select the folder for which you want to retrieve the file listing"
        Me.FolderBrowserDialog1.ShowNewFolderButton = False
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Black
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(464, 62)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Label4)
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(472, 96)
        Me.MinimumSize = New System.Drawing.Size(472, 96)
        Me.Name = "Main_Screen"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Single Level File Lister 1.0"
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub Error_Handler(ByVal ex As Exception)
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Monitoring Program encountered the following problem: " & vbCrLf & ex.ToString)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Single Level File Lister's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub



    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            dataloaded = True
            splash_loader.Visible = False
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub generatelisting(ByVal folder As String, ByVal savefile As String, ByVal optionselected As Integer)
        'optionselected: 0-filenames 1-filepaths
        Try
            Dim dinfo As DirectoryInfo = New DirectoryInfo(folder)
            If dinfo.Exists = True Then
                Dim files() As FileInfo
                Dim finfo As FileInfo
                files = dinfo.GetFiles
                Dim writer As StreamWriter = New StreamWriter(savefile, False)
                For Each finfo In files
                    Select Case optionselected
                        Case 0
                            writer.WriteLine(finfo.Name)
                        Case 1
                            writer.WriteLine(finfo.FullName)
                    End Select

                Next
                writer.Close()
                finfo = New FileInfo(savefile)
                If finfo.Exists Then
                    If finfo.Length > 0 Then
                        Printing(savefile)
                    End If
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Try
            Dim result As DialogResult = FolderBrowserDialog1.ShowDialog
            If result = DialogResult.OK Then
                result = SaveFileDialog1.ShowDialog
                If result = DialogResult.OK Then
                    Dim op As Option_Select = New Option_Select
                    result = op.ShowDialog
                    If result = DialogResult.OK Then
                        Dim optionselected As Integer = op.ComboBox1.SelectedIndex
                        generatelisting(FolderBrowserDialog1.SelectedPath, SaveFileDialog1.FileName, optionselected)
                        op.Dispose()
                    End If
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    ' The PrintPage event is raised for each page to be printed.
    Private Sub pd_PrintPage(ByVal sender As Object, ByVal ev As PrintPageEventArgs)
        Try

        
        Dim linesPerPage As Single = 0
        Dim yPos As Single = 0
        Dim count As Integer = 0
        Dim leftMargin As Single = ev.MarginBounds.Left
        Dim topMargin As Single = ev.MarginBounds.Top
        Dim line As String = Nothing

        ' Calculate the number of lines per page.
        linesPerPage = ev.MarginBounds.Height / printFont.GetHeight(ev.Graphics)

        ' Iterate over the file, printing each line.
        While count < linesPerPage
            line = streamToPrint.ReadLine()
            If line Is Nothing Then
                Exit While
            End If
            yPos = topMargin + count * printFont.GetHeight(ev.Graphics)
            ev.Graphics.DrawString(line, printFont, Brushes.Black, leftMargin, _
                yPos, New StringFormat)
            count += 1
        End While

        ' If more lines exist, print another page.
        If Not (line Is Nothing) Then
            ev.HasMorePages = True
        Else
            ev.HasMorePages = False
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    ' Print the file.
    Public Sub Printing(ByVal filepath As String)
        Try
            streamToPrint = New StreamReader(filepath)
            Try
                printFont = New Font("Arial", 10)
                Dim pd As New PrintDocument
                AddHandler pd.PrintPage, AddressOf pd_PrintPage
                ' Print the document.
                Dim preview As PrintDialog = New PrintDialog
                preview.Document = pd
                Dim result As DialogResult
                result = preview.ShowDialog
                If result = DialogResult.OK Then
                    pd.Print()
                End If

            Finally
                streamToPrint.Close()
            End Try
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub 'Printing    

End Class
