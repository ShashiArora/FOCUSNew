Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Configuration
Imports System.Collections
'Imports SoftBrands.FourthShift.Transaction
'Imports CrystalDecisions.CrystalReports.Engine
Imports System.Windows.Forms

Public Class UserAccess

    Inherits System.Windows.Forms.Form
    Private ConnectionString As String
    '   Public glbFSTIClient As FSTIClient
    'Public inv As Boolean
    'Public bNewInvoice As Boolean
    'Public headkey As Integer
    'Public linekey As Integer
    'Public iNo As Integer
    'Public pre As String
    'Public dNO As String
    'Public invoicecreate As Boolean
    'Public DUTYEXISTS As Boolean
    'Public mode As String
    'Public printtype As String
    'Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH102;Database=FSPrograms;User ID=fsadmin;Password=fsadm1n!"
    Public Shared ConnectionStringNew As String = "Server=TSSBLRFSH104;Database=FSPrograms;User ID=addonapp;Password=FS8.e$25"


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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtpwd As System.Windows.Forms.TextBox
    Friend WithEvents lblPwd As System.Windows.Forms.Label
    Friend WithEvents txtuserid As System.Windows.Forms.TextBox
    Friend WithEvents lblUserID As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnLogin As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(UserAccess))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnLogin = New System.Windows.Forms.Button()
        Me.txtpwd = New System.Windows.Forms.TextBox()
        Me.lblPwd = New System.Windows.Forms.Label()
        Me.txtuserid = New System.Windows.Forms.TextBox()
        Me.lblUserID = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Silver
        Me.GroupBox1.Controls.Add(Me.btnCancel)
        Me.GroupBox1.Controls.Add(Me.btnLogin)
        Me.GroupBox1.Controls.Add(Me.txtpwd)
        Me.GroupBox1.Controls.Add(Me.lblPwd)
        Me.GroupBox1.Controls.Add(Me.txtuserid)
        Me.GroupBox1.Controls.Add(Me.lblUserID)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.Black
        Me.GroupBox1.Location = New System.Drawing.Point(248, 16)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(240, 119)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Sign On"
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.Gray
        Me.btnCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.Color.Black
        Me.btnCancel.Location = New System.Drawing.Point(132, 83)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(64, 24)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnLogin
        '
        Me.btnLogin.BackColor = System.Drawing.Color.Gray
        Me.btnLogin.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLogin.ForeColor = System.Drawing.Color.Black
        Me.btnLogin.Location = New System.Drawing.Point(44, 83)
        Me.btnLogin.Name = "btnLogin"
        Me.btnLogin.Size = New System.Drawing.Size(64, 24)
        Me.btnLogin.TabIndex = 3
        Me.btnLogin.Text = "OK"
        Me.btnLogin.UseVisualStyleBackColor = False
        '
        'txtpwd
        '
        Me.txtpwd.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtpwd.Location = New System.Drawing.Point(96, 56)
        Me.txtpwd.MaxLength = 10
        Me.txtpwd.Name = "txtpwd"
        Me.txtpwd.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtpwd.Size = New System.Drawing.Size(100, 21)
        Me.txtpwd.TabIndex = 1
        '
        'lblPwd
        '
        Me.lblPwd.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPwd.Location = New System.Drawing.Point(24, 56)
        Me.lblPwd.Name = "lblPwd"
        Me.lblPwd.Size = New System.Drawing.Size(64, 23)
        Me.lblPwd.TabIndex = 6
        Me.lblPwd.Text = "Password"
        '
        'txtuserid
        '
        Me.txtuserid.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtuserid.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtuserid.Location = New System.Drawing.Point(96, 24)
        Me.txtuserid.MaxLength = 3
        Me.txtuserid.Name = "txtuserid"
        Me.txtuserid.Size = New System.Drawing.Size(100, 21)
        Me.txtuserid.TabIndex = 0
        '
        'lblUserID
        '
        Me.lblUserID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserID.Location = New System.Drawing.Point(24, 32)
        Me.lblUserID.Name = "lblUserID"
        Me.lblUserID.Size = New System.Drawing.Size(72, 16)
        Me.lblUserID.TabIndex = 4
        Me.lblUserID.Text = "User  Id"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(19, 14)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(223, 121)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 63
        Me.PictureBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.GroupBox1)
        Me.GroupBox2.Controls.Add(Me.PictureBox1)
        Me.GroupBox2.ForeColor = System.Drawing.Color.Black
        Me.GroupBox2.Location = New System.Drawing.Point(-2, -2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(500, 286)
        Me.GroupBox2.TabIndex = 64
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Login"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 200)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(216, 23)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "(c) 2013 Trelleborg.  All rights reserved"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(245, 192)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(272, 24)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Release : Ver 3.0  13Oct-2014"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(245, 247)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(176, 16)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Database : FSPrograms"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(245, 216)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(232, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Database Server: TSSBLRFSH101"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(245, 169)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(272, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Label1"
        '
        'UserAccess
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(501, 285)
        Me.Controls.Add(Me.GroupBox2)
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Name = "UserAccess"
        Me.Text = "Focus : Login  Screen"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub lblPwd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub lblUserID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtuserid_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtuserid.TextChanged


    End Sub

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click

        Dim cnSQL As SqlConnection = New SqlConnection(ConnectionStringNew)
        'Dim cnsql As SqlConnection = New sqlconnnection(System.Configuration.cnfigurationManager.connectionString("connection string"))

        'System.Configuration.cnfigurationManager.connectionString("connection string")






        Dim cmSQL As SqlCommand
        Dim drSQL As SqlDataReader
        Dim strSQL As String

        If txtuserid.Text = "" Then
            MsgBox("Enter User Name", vbInformation)
            Exit Sub

        End If

        If txtpwd.Text = "" Then
            MsgBox("Enter Password", vbInformation)
            Exit Sub

        End If

        strSQL = "Select User_Id,Password, AC_TYPE from dbo.ENQ_User_Access where ACTIVE_INACTIVE='A' AND User_Id = '" & txtuserid.Text & "'"

        cnSQL.Open()

        cmSQL = New SqlCommand(strSQL, cnSQL)
        drSQL = cmSQL.ExecuteReader()

        If drSQL.Read() Then
            If txtpwd.Text <> drSQL.Item(1) Then
                MsgBox("Invalid UserId or Password", vbInformation)
                Exit Sub
            End If
            username = txtuserid.Text
            usertype = Trim(drSQL.Item(2))

        Else
            MsgBox("Invalid UserId or Password", vbInformation)
            Exit Sub

        End If

        'Dim menu As New mainmenu

        username = txtuserid.Text

        Dim foc As New MainMenu
        'foc.Show()


        'inv.Show()
        ' Me.Close()
        Me.Hide()


    End Sub

    Private Sub UseAccess_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtuserid.Focus()
        Label1.Text = System.DateTime.Now()
        'Label6.Text = System.DayOfWeek.Friday





    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()

    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub txtuserid_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtuserid.KeyDown


    End Sub

    Private Sub txtuserid_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtuserid.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            txtpwd.Focus()

            Exit Sub

        End If

    End Sub

    Private Sub txtpwd_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtpwd.TextChanged

    End Sub

    Private Sub txtpwd_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtpwd.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            btnLogin.Focus()
            Exit Sub

        End If
    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub
End Class
