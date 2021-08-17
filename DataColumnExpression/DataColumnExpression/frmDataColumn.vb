#Region "ABOUT"
' / --------------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / MORE: http://www.g2gnet.com/webboard
' /
' / Purpose: Data Column Expression with VB.NET (2010).
' / Microsoft Visual Basic .NET (2010) + MS Access 2003+
' /
' / This is open source code under @CopyLeft by Thongkorn/Common Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / --------------------------------------------------------------------------
#End Region

Imports System.Data.OleDb

Public Class frmDataColumn

    ' / --------------------------------------------------------------------------
    Private Sub frmDataColumn_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        '// https://docs.microsoft.com/en-us/dotnet/api/system.data.datacolumn.expression?view=netframework-4.7.2
        '// Gets or sets the expression used to filter rows, calculate the values in a column, or create an aggregate column.
        Call DataColumnExpressions()
    End Sub

    ' / --------------------------------------------------------------------------
    Sub DataColumnExpressions()
        Dim Conn As New System.Data.OleDb.OleDbConnection
        Dim DA As New System.Data.OleDb.OleDbDataAdapter
        Dim Cmd As New System.Data.OleDb.OleDbCommand
        Dim DC As DataColumn
        Dim DS As New DataSet
        '// Connect to the database
        Conn.ConnectionString = _
            " Provider=Microsoft.ACE.OLEDB.12.0; " & _
            " Data Source=" & _
            MyPath(Application.StartupPath) & "SampleDB.accdb;"
        Cmd.Connection = Conn
        DA.SelectCommand = Cmd

        '// Fill the DataTable
        Cmd.CommandText = "SELECT * FROM Products"
        DA.Fill(DS, "Products")

        '// Add on a few simple expression columns
        DC = New DataColumn("UnitPriceX2")
        DC.DataType = GetType(Double)
        DC.Expression = "UnitPrice * 2" '/ Column and numeric constant.
        DS.Tables("Products").Columns.Add(DC)

        '// Add on a few simple expression columns
        DC = New DataColumn("UnitPriceX2X2")
        DC.DataType = GetType(Double)
        DC.Expression = "UnitPriceX2 * 2" '/ Column and numeric constant.
        DS.Tables("Products").Columns.Add(DC)

        DC = New DataColumn("UnitsInStockUpdate")
        DC.DataType = GetType(Integer)
        DC.Expression = "UnitsInStock" '/ Column UnitsInStock.
        DS.Tables("Products").Columns.Add(DC)

        DC = New DataColumn("UnitsOnOrderUpdate")
        DC.DataType = GetType(Integer)
        DC.Expression = "UnitsOnOrder" '/ Column UnitsOnOrder.
        DS.Tables("Products").Columns.Add(DC)

        '// Summary expression UnitsInStock + UnitsOnOrder.
        DC = New DataColumn("TotalStock")
        DC.DataType = GetType(Integer)
        DC.Expression = "UnitsInStock + UnitsOnOrder" '/ Add two columns.
        DS.Tables("Products").Columns.Add(DC)

        '// Text or String.
        'DC = New DataColumn("New Header")
        'DC.DataType = GetType(String)
        'DC.Expression = "ProductName + ' New Header'" '/ Column and string constant
        'DS.Tables("Products").Columns.Add(DC)
        '//
        dgvData.DataSource = DS.Tables("Products").DefaultView
        '// Organized DataGridView.
        Call SetupDataGridView(dgvData)
        '//
        DS.Dispose()
        DC.Dispose()
        DA.Dispose()
        Cmd.Dispose()
        Conn.Close()
    End Sub

    ' / --------------------------------------------------------------------------
    ' / Organized DataGridView.
    Private Sub SetupDataGridView(ByRef DGV As DataGridView)
        With DGV
            .Columns.Clear()
            .RowTemplate.Height = 26
            .AllowUserToOrderColumns = True
            .AllowUserToDeleteRows = False
            .AllowUserToAddRows = False
            .ReadOnly = True
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .Font = New Font("Tahoma", 8)
            .AlternatingRowsDefaultCellStyle.BackColor = Color.OldLace
            .DefaultCellStyle.SelectionBackColor = Color.SeaGreen
            '.AlternatingRowsDefaultCellStyle.BackColor = Color.LightYellow
            '.DefaultCellStyle.SelectionBackColor = Color.LightBlue
            '/ Auto size column width of each main by sorting the field.
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            .AutoResizeColumns()

            Dim ColumnID As New DataGridViewTextBoxColumn
            With ColumnID
                .DataPropertyName = "ProductID"
                .Name = "ProductID"
                .HeaderText = "ProductID"
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            .Columns.Add(ColumnID)

            Dim Column1 As New DataGridViewTextBoxColumn
            With Column1
                .DataPropertyName = "ProductName"
                .Name = "ProductName"
                .HeaderText = "ProductName"
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            .Columns.Add(Column1)

            Dim Column2 As New DataGridViewTextBoxColumn
            With Column2
                .DataPropertyName = "UnitPrice"
                .Name = "UnitPrice"
                .HeaderText = "UnitPrice"
                .DefaultCellStyle.Format = "N2"
            End With
            .Columns.Add(Column2)

            Dim Column3 As New DataGridViewTextBoxColumn
            With Column3
                .DataPropertyName = "UnitPriceX2"
                .Name = "UnitPriceX2"
                .HeaderText = "UnitPrice x 2"
                .DefaultCellStyle.Format = "N2"
            End With
            .Columns.Add(Column3)

            Dim UnitPriceX4 As New DataGridViewTextBoxColumn
            With UnitPriceX4
                .DataPropertyName = "UnitPriceX2X2"
                .Name = "UnitPriceX2X2"
                .HeaderText = "UnitPriceX2 x 2"
                .DefaultCellStyle.Format = "N2"
            End With
            .Columns.Add(UnitPriceX4)

            Dim Column4 As New DataGridViewTextBoxColumn
            With Column4
                .DataPropertyName = "UnitsInStockUpdate"
                .Name = "UnitsInStock"
                .HeaderText = "Units In Stock"
                .DefaultCellStyle.Format = "N2"
            End With
            .Columns.Add(Column4)

            Dim Column5 As New DataGridViewTextBoxColumn
            With Column5
                .DataPropertyName = "UnitsOnOrderUpdate"
                .Name = "UnitsOnOrder"
                .HeaderText = "Units On Order"
                .DefaultCellStyle.Format = "N2"
            End With
            .Columns.Add(Column5)

            Dim Column6 As New DataGridViewTextBoxColumn
            With Column6
                .DataPropertyName = "TotalStock"
                .Name = "TotalStock"
                .HeaderText = "Total Stock"
                .DefaultCellStyle.Format = "N2"
            End With
            .Columns.Add(Column6)

            '// For Column and string constant.
            'Dim Column7 As New DataGridViewTextBoxColumn
            'With Column7
            '.DataPropertyName = "New Header"
            '.Name = "New Header"
            '.HeaderText = "New Header"
            '.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            'End With
            '.Columns.Add(Column7)
        End With

        '// Columns Alignment.
        For i = 2 To 6
            With DGV.Columns(i)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
            End With
        Next
    End Sub

    Private Sub frmDataColumn_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
        GC.SuppressFinalize(Me)
        Application.Exit()
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Get my project path
    ' / AppPath = C:\My Project\bin\debug
    ' / Replace "\bin\debug" with ""
    ' / Return : C:\My Project\
    Function MyPath(AppPath As String) As String
        '/ MessageBox.Show(AppPath);
        AppPath = AppPath.ToLower()
        '/ Return Value
        MyPath = AppPath.Replace("\bin\debug", "").Replace("\bin\release", "").Replace("\bin\x86\debug", "")
        '// If not found folder then put the \ (BackSlash) at the end.
        If Microsoft.VisualBasic.Right(MyPath, 1) <> Chr(92) Then MyPath = MyPath & Chr(92)
    End Function
End Class
