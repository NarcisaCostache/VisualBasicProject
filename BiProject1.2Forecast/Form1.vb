Imports Microsoft.Office.Interop
Imports System.Data.SqlClient
Public Class Form1
    Dim conn_string = "server= LAPTOP-LB4U4SD0\SQLEXPRESS;database= BiProject; integrated security= sspi"
    Dim query = "select left(convert(varchar(10), SaleDate, 111),7) as SalesYearMonth, " _
    & "avg(TotalPrice) as Sales " _
  & "from [BiProject].[dbo].[Sales]"

    Dim lc = " where SaleDate>="
    Dim rc = " AND SaleDate<="
    Dim goby = " group by left(convert(varchar(10), SaleDate,111),7) " _
    & "order by left(convert(varchar(10), SaleDate,111),7)"
    Dim table = "WineAverageSales"

    ''' 




    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim conn_str = "server= LAPTOP-LB4U4SD0\SQLEXPRESS;database= BiProject; integrated security= sspi"
        Dim conn = New SqlConnection(conn_str)
        conn.Open()
        Dim q = "Select month(SaleDate)As month , avg(TotalPrice) As AverageSales from [BiProject].[dbo].[Sales]group by month(SaleDate) order by month(SaleDate)"
        Dim da = New SqlDataAdapter(q, conn)
        Dim ds = New DataSet
        Dim tbl = "WineAverageSales"
        da.Fill(ds, tbl)
        DataGridView1.DataMember = tbl
        DataGridView1.DataSource = ds
        For i = 0 To ds.Tables(tbl).Rows.Count - 1
            Dim oyv = ds.Tables(tbl).Rows(i).Item("month")
            Dim asv = ds.Tables(tbl).Rows(i).Item("AverageSales")
            Chart1.Series("AverageSales").Points.AddXY(oyv, asv)
        Next

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim conn_str = "server= LAPTOP-LB4U4SD0\SQLEXPRESS;database= BiProject; integrated security= sspi"
        Dim conn = New SqlConnection(conn_str)
        conn.Open()
        Dim q = "Select month(SaleDate)As month , avg(TotalPrice) As AverageSales from [BiProject].[dbo].[Sales]group by month(SaleDate) order by month(SaleDate)"
        Dim da = New SqlDataAdapter(q, conn)
        Dim ds = New DataSet
        Dim tbl = "WineAverageSales"
        da.Fill(ds, tbl)
        Dim n = ds.Tables(tbl).Rows.Count
        Dim x(n) As Double
        Dim y(n) As Double
        Dim sum_of_xy As Double = 0
        Dim sum_of_xx As Double = 0
        Dim sum_of_x As Double = 0
        Dim sum_of_y As Double = 0
        For i = 0 To n - 1
            x(i) = i + 1
            y(i) = ds.Tables(tbl).Rows(i).Item("AverageSales")
            sum_of_xy = sum_of_xy + (x(i) * y(i))
            sum_of_xx = sum_of_xx + (x(i) * x(i))
            sum_of_x = sum_of_x * x(i)
            sum_of_y = sum_of_y * y(i)

        Next
        Dim alpha As Double
        alpha = ((n * sum_of_xy) - (sum_of_x * sum_of_y)) / ((n * sum_of_xx) - (sum_of_x * sum_of_x))
        Dim beta As Double
        beta = (sum_of_y - (alpha * sum_of_x)) / n
        Dim month1 = ds.Tables(tbl).Rows(0).Item("month")

        Dim input_x = (TextBox1.Text - month1 + 1)
        Dim pred_y As Double
        pred_y = (alpha * input_x) + beta
        Chart1.Series("AverageSales").Points.AddXY((input_x + month1 - 1), pred_y)
        TextBox1.Text = pred_y.ToString



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim lq = query _
        & lc & "'" & DateTimePicker1.Value.ToShortDateString.ToString() & "'" _
        & rc & "'" & DateTimePicker2.Value.ToShortDateString.ToString() & "'" _
        & goby
        ' Dim conn_str = "server= LAPTOP-LB4U4SD0\SQLEXPRESS;database= BiProject; integrated security= sspi"

        Dim conn1 = New SqlConnection(conn_string)
        conn1.Open()
        Dim da1 = New SqlDataAdapter(lq, conn1)
        Dim ds1 = New DataSet
        da1.Fill(ds1, table)
        DataGridView1.DataMember = table
        DataGridView1.DataSource = ds1
        Dim n = ds1.Tables(table).Rows.Count - 1
        Dim x
        Dim y
        For Each cs As DataVisualization.Charting.Series In Chart1.Series
            cs.Points.Clear()
        Next
        For i = 0 To n
            x = ds1.Tables(table).Rows(i).Item("SalesYearMonth")
            y = ds1.Tables(table).Rows(i).Item("Sales")
            Chart1.Series("AverageSales").Points.AddXY(x, y)
        Next

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim conn_string2 = "server= LAPTOP-LB4U4SD0\SQLEXPRESS;database= BiProject; integrated security= sspi"
        Dim query2 = "select cast(year(SaleDate) as char(4))+'-'+cast(month(SaleDate) 
  as char(2))as ordym, avg(TotalPrice) as Vanzari 
  from [BiProject].[dbo].[Sales]
  group by year(SaleDate), month(SaleDate)
  order by year(SaleDate), month(SaleDate)"

        Dim table2 = "WineAverageSales2"
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
            matr(i, 0) = ds2.Tables(table2).Rows(i).Item("ordym")
            matr(i, 1) = ds2.Tables(table2).Rows(i).Item("Vanzari")

        Next
        oSheet.Range("A1").Resize(n + 1, 2).Value = matr
        oCh = oSheet.ChartObjects.Add(150, 40, 800, 400).Chart
        oCh.ChartType = Excel.XlChartType.xlLine
        oCh.SetSourceData(oSheet.Range("A1:B" & (n + 1).ToString))
        oCh.SeriesCollection(1).name = "AverageSales"
        Dim cboSel = Me.ComboBox1.SelectedItem.ToString
        Dim cboVal = Me.TextBox2.Text
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

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        Dim conn_string2 = "server= LAPTOP-LB4U4SD0\SQLEXPRESS;database= BiProject; integrated security= sspi"
        Dim query2 = "select cast(year(SaleDate) as char(4))+'-'+cast(month(SaleDate) 
  as char(2))as ordym, avg(TotalPrice) as Vanzari 
  from [BiProject].[dbo].[Sales]
  group by year(SaleDate), month(SaleDate)
  order by year(SaleDate), month(SaleDate)"
        Dim table2 = "WineAverageSales2"
        Dim conn2 = New SqlConnection(conn_string2)
        conn2.Open()
        Dim da3 = New SqlDataAdapter(query2, conn2)
        Dim ds3 = New DataSet
        da3.Fill(ds3, table2)
        Dim oXl As Object
        Dim oBook As Object
        Dim oSheet As Object
        Dim oCh As Object
        oXl = CreateObject("Excel.application")
        oBook = oXl.Workbooks.Add
        oSheet = oBook.Worksheets.Item(1)
        Dim n = ds3.Tables(table2).Rows.Count - 1
        Dim matr(n, 1)
        For i = 0 To n
            matr(i, 0) = ds3.Tables(table2).Rows(i).Item("ordym")
            matr(i, 1) = ds3.Tables(table2).Rows(i).Item("Vanzari")

        Next
        Dim trend(8, 1) As String
        trend(0, 0) = "4133"
        trend(0, 1) = "Logarithmic"
        trend(1, 0) = "4132"
        trend(1, 1) = "Linear"
        For k = 2 To 6
            trend(k, 0) = "3"
            trend(k, 1) = "Polynomial" & k.ToString

        Next
        trend(7, 0) = "4"
        trend(7, 1) = "Power"
        trend(8, 0) = "5"
        trend(8, 1) = "Exponential"
        Dim o As Integer
        Dim max_R_sq = ""
        Dim mem_type As Integer
        For k = 0 To 6
            oSheet = oBook.Worksheets.Add
            oSheet.Range("A1").Resize(n + 1, 2).Value = matr
            oCh = oSheet.ChartObjects.Add(150, 40, 800, 400).Chart
            oCh.ChartType = Excel.XlChartType.xlXYScatterLines
            oCh.SetSourceData(oSheet.Range("A1:B" & (n + 1).ToString))
            oCh.SeriesCollection(1).name = "AverageSales"
            If k >= 2 And k <= 6 Then
                o = Int32.Parse(Mid(trend(k, 1), 11, 1))
                oCh.SeriesCollection(1).TrendLines.Add(Type:=trend(k, 0), order:=o, Forward:=1, Backward:=0, DisplayEquation:=False, DisplayRSquared:=False).select()
            Else
                oCh.SeriesCollection(1).TrendLines.Add(Type:=trend(k, .0), Foreward:=1, Backward:=0, DisplayEquation:=False, DisplayRSquared:=True).Select()

            End If
            Dim R_Square = oCh.SeriesCollection(1).TrendLines(1).DataLable.text
            R_Square = Mid(R_Square, 6, Len(R_Square) - 5)
            If R_Square > max_R_sq Then
                max_R_sq = R_Square
                mem_type = k

            End If
        Next
        MsgBox("Your data indicate the " & trend(mem_type, 1).ToString & " trend from those 5 on the upper left, with a max of: " & max_R_sq)
        oBook.Close(savechanges:=False)
        oXl.Quit()


    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub
End Class
