Imports System.IO

Public Class Option_Select
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Public WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Option_Select))
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Button6 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ComboBox1
        '
        Me.ComboBox1.Items.AddRange(New Object() {"File Names Only", "Full File Paths"})
        Me.ComboBox1.Location = New System.Drawing.Point(16, 16)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(176, 21)
        Me.ComboBox1.TabIndex = 0
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.Black
        Me.Button6.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button6.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button6.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.ForeColor = System.Drawing.Color.Yellow
        Me.Button6.Location = New System.Drawing.Point(200, 16)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(64, 24)
        Me.Button6.TabIndex = 11
        Me.Button6.Text = "OK"
        '
        'Option_Select
        '
        Me.AcceptButton = Me.Button6
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Yellow
        Me.ClientSize = New System.Drawing.Size(282, 56)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.ComboBox1)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(288, 80)
        Me.MinimumSize = New System.Drawing.Size(288, 80)
        Me.Name = "Option_Select"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Option Select"
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

    Private Sub Option_Select_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            ComboBox1.SelectedIndex = 0
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Try
            Me.Hide()
        Catch ex As Exception
            Error_Handler(ex)
        End Try

    End Sub
End Class
