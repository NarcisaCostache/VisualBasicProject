Imports Microsoft.Office.Interop
Imports System.Data.SqlClient
Public Class Form2
    Dim conn_string2 = "server= LAPTOP-LB4U4SD0\SQLEXPRESS;database= BiProject; integrated security= sspi"
    Dim query2 = " select cast(year(SaleDate) as char(4))+'-'+cast(month(SaleDate) 
  as char(2))as ordym," _
 & "avg(TotalPrice) as Vanzari" _
  & "from [BiProject].[dbo].[Sales]" _
  & "group by year(SaleDate), month(SaleDate)" _
 & " order by year(SaleDate), month(SaleDate)"
    Dim table2 = "WineAverageSales"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim conn2 = New SqlConnection(conn_string2)
        conn2.Open()
        Dim da2 = New SqlDataAdapter(query2, conn2)
        Dim ds2 = New DataSet
        da2.Fill(ds2, table2)
        Dim oXl As Object
        Dim oBook As Object
        Dim oSheet As Object
        Dim oCh As Object
        oXl = CreateObject("Excel.application")
        oBook = oXl.Workbooks.Add
        oSheet = oBook.Worksheets.Item(1)
        Dim n = ds2.Tables(table2).Rows.Count - 1
        Dim matr(n, 1)
        For i = 0 To n
            matr(i, 0) = ds2.Tables(table2).Rows(i).Item("SalesYearMonth")
            matr(i, 1) = ds2.Tables(table2).Rows(i).Item("Sales")

        Next
        oSheet.Range("A1").Resize(n + 1, 2).Value = matr
        oCh = oSheet.ChartObjects.Add(150, 40, 800, 400).Chart
        oCh.ChartType = Excel.XlChartType.xlLine
        oCh.SetSourceData(oSheet.Range("A1:B" & (n + 1).ToString))
        oCh.SeriesCollection(1).name = "AverageSales"
        Dim cboSel = Me.ComboBox1.SelectedItem.ToString
        Dim cboVal = Me.TextBox1.Text
        Dim op As Integer
        op = Int32.Parse(Me.ComboBox1.SelectedItem)
        oCh.SeriesCollection(1).Trendlines.Add(Type:=cboSel, Forward:=cboVal, Backward:=0, DisplayEquation:=True, DisplayRSquared:=False).select()
        Dim R_square = oCh.SeriesCollection(1).TrendLines(1).DataLabel.text
        For i = 0 To n
            oSheet.Range("C" & (i + 1))(i + 1).ToString()
        Next
        oSheet.Range("D1") = R_square
        oSheet.Range("E1") = Mid(R_square, 6, Len(R_square) - 5)
        oCh.SeriesCollection(1).TrendLines(1).DisplayRSquared = False
        oCh.SeriesCollection(1).TrendLines(1).DisplayEquation = True
        Dim T_eq = oCh.SeriesCollection(1).TrendLines(1).DataLabel.text
        oSheet.Range("F1") = T_eq
        oXl.Visible = True
        oXl.UserControl = True




    End Sub
End Class