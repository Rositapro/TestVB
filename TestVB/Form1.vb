Imports OfficeOpenXml

Public Class Form1

    Dim id As Integer

    Public Sub New()
        InitializeComponent()
        id = 0
        lstvData.FullRowSelect = True

        ' Establecer el contexto de la licencia de EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial ' o LicenseContext.Commercial dependiendo de tu caso
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim x As MyItem
        Dim description As String = txtDescription.Text
        id = id + 1
        Dim Price As Random = New Random()

        x = New MyItem(id, description, Math.Round(Price.NextDouble() * 1000, 2))
        lstItems.Items.Add(x.ToString())

        For i = 1 To 100
            Dim row As ListViewItem = New ListViewItem(x.Id)
            row.SubItems.Add(x.Description)
            row.SubItems.Add(x.Price)
            lstvData.Items.Add(row)
            x.Id = x.Id + 1
            x.Price = Math.Round(Price.NextDouble() * 1000, 2)
        Next

        UpdateLabel()
        UpdateTotal()
    End Sub

    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        If lstvData.SelectedItems.Count = 0 Then
            Return
        End If
        For Each item As ListViewItem In lstvData.SelectedItems
            lstvData.Items.Remove(item)
        Next
        UpdateLabel()
        UpdateTotal()
    End Sub

    Sub UpdateLabel()
        lblCount.Text = lstvData.Items.Count
    End Sub

    Sub UpdateTotal()
        Dim Total As Decimal = 0
        For Each item As ListViewItem In lstvData.Items
            Total = Total + Decimal.Parse(item.SubItems(2).Text)
        Next
        LblTotal.Text = Total
    End Sub

    Private Sub btnSaveExcel_Click(sender As Object, e As EventArgs) Handles btnSaveExcel.Click
        Using package As New ExcelPackage()
            Dim worksheet = package.Workbook.Worksheets.Add("Sheet1")

            ' Headers
            worksheet.Cells("A1").Value = "ID"
            worksheet.Cells("B1").Value = "Descripción"
            worksheet.Cells("C1").Value = "Precio"

            ' Data
            For i = 0 To lstvData.Items.Count - 1
                worksheet.Cells(i + 2, 1).Value = lstvData.Items(i).SubItems(0).Text ' ID
                worksheet.Cells(i + 2, 2).Value = lstvData.Items(i).SubItems(1).Text ' Descripción
                worksheet.Cells(i + 2, 3).Value = lstvData.Items(i).SubItems(2).Text ' Precio
            Next

            ' Save the workbook
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Excel Files|*.xlsx"
            saveFileDialog.FileName = "ExcelData.xlsx"
            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Dim file As New IO.FileInfo(saveFileDialog.FileName)
                package.SaveAs(file)
                MessageBox.Show("Datos guardados en " & file.FullName, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Using
    End Sub

End Class




