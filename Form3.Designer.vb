<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim ReportDataSource1 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form3))
        Me.TRANSFER2014BindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.INTERSWITCHDataSet = New Interswitch_Automation.INTERSWITCHDataSet()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.ReportViewer1 = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.TRANSFER2014TableAdapter = New Interswitch_Automation.INTERSWITCHDataSetTableAdapters.TRANSFER2014TableAdapter()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        CType(Me.TRANSFER2014BindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.INTERSWITCHDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TRANSFER2014BindingSource
        '
        Me.TRANSFER2014BindingSource.DataMember = "TRANSFER2014"
        Me.TRANSFER2014BindingSource.DataSource = Me.INTERSWITCHDataSet
        '
        'INTERSWITCHDataSet
        '
        Me.INTERSWITCHDataSet.DataSetName = "INTERSWITCHDataSet"
        Me.INTERSWITCHDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(534, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(141, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Upload Computation File"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(534, 72)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(141, 23)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Delete Computation File"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(534, 101)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(141, 23)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "Next"
        Me.Button3.UseVisualStyleBackColor = True
        Me.Button3.Visible = False
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Location = New System.Drawing.Point(18, 15)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(485, 20)
        Me.DateTimePicker1.TabIndex = 3
        '
        'ReportViewer1
        '
        Me.ReportViewer1.BackColor = System.Drawing.Color.Honeydew
        ReportDataSource1.Name = "DataSet1"
        ReportDataSource1.Value = Me.TRANSFER2014BindingSource
        Me.ReportViewer1.LocalReport.DataSources.Add(ReportDataSource1)
        Me.ReportViewer1.LocalReport.ReportEmbeddedResource = "Interswitch_Automation.Report2.rdlc"
        Me.ReportViewer1.Location = New System.Drawing.Point(18, 162)
        Me.ReportViewer1.Name = "ReportViewer1"
        Me.ReportViewer1.Size = New System.Drawing.Size(657, 494)
        Me.ReportViewer1.TabIndex = 4
        Me.ReportViewer1.Visible = False
        '
        'TRANSFER2014TableAdapter
        '
        Me.TRANSFER2014TableAdapter.ClearBeforeFill = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(534, 41)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(141, 23)
        Me.Button4.TabIndex = 5
        Me.Button4.Text = "View Report"
        Me.Button4.UseVisualStyleBackColor = True
        Me.Button4.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox1.Location = New System.Drawing.Point(18, 47)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(484, 99)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.ClientSize = New System.Drawing.Size(695, 676)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.ReportViewer1)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Form3"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "INTERSWITCH AUTOMATION"
        CType(Me.TRANSFER2014BindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.INTERSWITCHDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents ReportViewer1 As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents TRANSFER2014BindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents INTERSWITCHDataSet As Interswitch_Automation.INTERSWITCHDataSet
    Friend WithEvents TRANSFER2014TableAdapter As Interswitch_Automation.INTERSWITCHDataSetTableAdapters.TRANSFER2014TableAdapter
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
End Class
