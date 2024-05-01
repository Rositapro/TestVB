<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        btnSaveExcel = New Button()
        LblTotal = New Label()
        btnRemove = New Button()
        lblCount = New Label()
        lstvData = New ListView()
        ColumnHeader1 = New ColumnHeader()
        ColumnHeader2 = New ColumnHeader()
        ColumnHeader3 = New ColumnHeader()
        lstItems = New ListBox()
        btnSave = New Button()
        txtDescription = New TextBox()
        Label1 = New Label()
        SuspendLayout()
        ' 
        ' btnSaveExcel
        ' 
        btnSaveExcel.Font = New Font("Segoe UI", 15F)
        btnSaveExcel.Location = New Point(630, 50)
        btnSaveExcel.Name = "btnSaveExcel"
        btnSaveExcel.Size = New Size(139, 38)
        btnSaveExcel.TabIndex = 17
        btnSaveExcel.Text = "Save"
        btnSaveExcel.UseVisualStyleBackColor = True
        ' 
        ' LblTotal
        ' 
        LblTotal.AutoSize = True
        LblTotal.Font = New Font("Segoe UI", 15F)
        LblTotal.Location = New Point(575, 376)
        LblTotal.Name = "LblTotal"
        LblTotal.Size = New Size(0, 28)
        LblTotal.TabIndex = 16
        ' 
        ' btnRemove
        ' 
        btnRemove.Font = New Font("Segoe UI", 15F)
        btnRemove.Location = New Point(485, 50)
        btnRemove.Name = "btnRemove"
        btnRemove.Size = New Size(139, 38)
        btnRemove.TabIndex = 15
        btnRemove.Text = "Remove"
        btnRemove.UseVisualStyleBackColor = True
        ' 
        ' lblCount
        ' 
        lblCount.AutoSize = True
        lblCount.Font = New Font("Segoe UI", 15F)
        lblCount.Location = New Point(322, 357)
        lblCount.Name = "lblCount"
        lblCount.Size = New Size(0, 28)
        lblCount.TabIndex = 14
        ' 
        ' lstvData
        ' 
        lstvData.Columns.AddRange(New ColumnHeader() {ColumnHeader1, ColumnHeader2, ColumnHeader3})
        lstvData.Font = New Font("Segoe UI", 15F)
        lstvData.Location = New Point(322, 143)
        lstvData.Name = "lstvData"
        lstvData.Size = New Size(415, 200)
        lstvData.TabIndex = 13
        lstvData.UseCompatibleStateImageBehavior = False
        lstvData.View = View.Details
        ' 
        ' ColumnHeader1
        ' 
        ColumnHeader1.Text = "ID"
        ' 
        ' ColumnHeader2
        ' 
        ColumnHeader2.Text = "DESCRIPTION"
        ColumnHeader2.Width = 200
        ' 
        ' ColumnHeader3
        ' 
        ColumnHeader3.Text = "PRICE"
        ColumnHeader3.Width = 120
        ' 
        ' lstItems
        ' 
        lstItems.Font = New Font("Segoe UI", 15F)
        lstItems.FormattingEnabled = True
        lstItems.ItemHeight = 28
        lstItems.Location = New Point(32, 143)
        lstItems.Name = "lstItems"
        lstItems.Size = New Size(248, 200)
        lstItems.TabIndex = 12
        ' 
        ' btnSave
        ' 
        btnSave.Font = New Font("Segoe UI", 15F)
        btnSave.Location = New Point(340, 50)
        btnSave.Name = "btnSave"
        btnSave.Size = New Size(139, 38)
        btnSave.TabIndex = 11
        btnSave.Text = "Save"
        btnSave.UseVisualStyleBackColor = True
        ' 
        ' txtDescription
        ' 
        txtDescription.Font = New Font("Segoe UI", 15F)
        txtDescription.Location = New Point(150, 47)
        txtDescription.Name = "txtDescription"
        txtDescription.Size = New Size(164, 34)
        txtDescription.TabIndex = 10
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Font = New Font("Segoe UI", 15F)
        Label1.Location = New Point(32, 47)
        Label1.Name = "Label1"
        Label1.Size = New Size(112, 28)
        Label1.TabIndex = 9
        Label1.Text = "Description"
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 450)
        Controls.Add(btnSaveExcel)
        Controls.Add(LblTotal)
        Controls.Add(btnRemove)
        Controls.Add(lblCount)
        Controls.Add(lstvData)
        Controls.Add(lstItems)
        Controls.Add(btnSave)
        Controls.Add(txtDescription)
        Controls.Add(Label1)
        Name = "Form1"
        Text = "Form1"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents btnSaveExcel As Button
    Friend WithEvents LblTotal As Label
    Friend WithEvents btnRemove As Button
    Friend WithEvents lblCount As Label
    Friend WithEvents lstvData As ListView
    Friend WithEvents ColumnHeader1 As ColumnHeader
    Friend WithEvents ColumnHeader2 As ColumnHeader
    Friend WithEvents ColumnHeader3 As ColumnHeader
    Friend WithEvents lstItems As ListBox
    Friend WithEvents btnSave As Button
    Friend WithEvents txtDescription As TextBox
    Friend WithEvents Label1 As Label

End Class
