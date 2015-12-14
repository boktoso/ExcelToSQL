Imports WindowsApplication1.Microsoft.Office.Interop
Imports WindowsApplication1.Microsoft.Office.Interop.Excel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Data.OleDb

Public Class ExcelToSQL
    ' ADD FEATURE: Allow user to select the columns in the Excel Speadsheet to use
    Public ExcelFilePath As String
    Public row_Num As Integer
    Public column_Num As Integer
    Public col_Max As Integer
    Public SQLFilePath As String
    Public con As New System.Data.OleDb.OleDbConnection
    Public cmd As New System.Data.OleDb.OleDbCommand
    Public ExcelArray(50) As String
    Public ExcelNumber As Integer = 0
    Public CurrentExcelNum As Integer = 0
    Public start_Row As Integer
    Public start_Column As Integer
    Public table As String
    Public database As String
    Public server As String

    'All Declarations of Classes

    Public Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles row_Num_Change.ValueChanged
        row_Num = row_Num_Change.Value
        start_Row = row_Num_Change.Value
        'Sets Row field to respected Row Integers
    End Sub

    Public Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles column_Num_Change.ValueChanged
        column_Num = column_Num_Change.Value
        start_Column = column_Num_Change.Value
        'Set Column field to respected Column Integers
    End Sub

    Public Sub col_Num_Max_ValueChanged(sender As Object, e As EventArgs) Handles col_Num_Max.ValueChanged
        col_Max = col_Num_Max.Value
    End Sub

    Public Sub Compile_Click(sender As Object, e As EventArgs) Handles b_Compile.Click
        Dim value As String
        While CurrentExcelNum < ExcelNumber 'continues through the Array until the last selected number (to avoid null errors)
            value = ExcelArray(CurrentExcelNum)
            TestResults.Text = "Test In Process"
            Dim Aggregate As New Aggregator()
            Aggregate.ExcelDump(value)
            CurrentExcelNum += 1
            'While there is a spreadsheet to integrate, grab table and dump into databse
        End While
        ExcelList.Items.Clear() 'clears out ExcelList field so no duplicate tables are copied by accident
        Array.Clear(ExcelArray, 0, ExcelArray.Length)
        ExcelNumber = 0
        TestResults.Text = "Test Completed"
    End Sub

    Public Sub ExcelFile_Click(sender As System.Object, e As System.EventArgs) Handles b_ExcelBrowse.Click
        Dim fDialog As New OpenFileDialog
        Dim fName As String
        fDialog.Filter = "xls files (*.xls)|*.xls|xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
        fDialog.FilterIndex = 2
        fDialog.RestoreDirectory = True
        fDialog.Multiselect = True
        If fDialog.ShowDialog() = DialogResult.OK Then
            For Each fName In fDialog.FileNames
                ExcelList.Items.Add(fName)
                ExcelArray(ExcelNumber) = fName
                ExcelNumber += 1
            Next
        End If
        'GETS Excel File(s) and checks for validation on correct .xls or .xlsx extension
    End Sub

    Public Sub b_ClearList_Click(sender As Object, e As EventArgs) Handles b_ClearList.Click
        ExcelList.Items.Clear()
        Array.Clear(ExcelArray, 0, ExcelArray.Length)
        ExcelNumber = 0
    End Sub

    Public Sub SQLConnect_Click(sender As System.Object, e As System.EventArgs) Handles b_SQLBrowse.Click
        'Server IP Address 127.0.0.1
        'Database is Database1
        'ID=ExcelAggregator
        'Pass=ExcelAggregator
        server = tb_SQLServer.Text
        database = tb_SQLDatabase.Text
        Try
            Dim connString As String = "Provider=SQLOLEDB;Server=@SERVER;Initial Catalog=@DATABASE;User Id=ExcelAggregator;Password=ExcelAggregator;"
            con = New OleDb.OleDbConnection(connString)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@SERVER", server)
            cmd.Parameters.AddWithValue("@DATABASE", database)
            con.Open()
            MessageBox.Show("Connected Successfully")
        Catch ex As Exception
            MessageBox.Show("Error While Connecting: " & ex.Message)
        Finally

        End Try
    End Sub

    Private Sub ExcelToSQL_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.FormClosing
        con.Close() 'ends connection to database once the form is closed
    End Sub

End Class

Public Class Aggregator
    Dim row_Num As Integer = ExcelToSQL.row_Num
    Dim column_Num As Integer = ExcelToSQL.column_Num
    Dim start_Column As Integer = ExcelToSQL.start_Column
    Dim start_Row As Integer = ExcelToSQL.start_Row
    Dim col_Max As Integer = ExcelToSQL.col_Max
    Dim content(25) As String
    Dim APP As Excel.ApplicationClass
    Dim worksheet As Excel.Worksheet
    Dim workbook As Excel.Workbook
    Dim total_Records As Integer = 0
    Dim failed_Records As Integer = 0
    Dim succeeded_Records As Integer
    Dim column_Name(25) As String
    Dim column_Count As Integer = 0
    Dim cmd As New System.Data.OleDb.OleDbCommand
    Dim table As String = ExcelToSQL.table
    Dim database As String = ExcelToSQL.database
    Dim server As String = ExcelToSQL.server

    Public Sub Get_Excel(ByVal value As String)
        APP = CreateObject("Excel.Application")
        workbook = APP.Workbooks.Open(value)
        worksheet = workbook.Worksheets(1)
        'Opens the Excel Spreadsheet
    End Sub

    Public Sub ExcelDump(ByVal value As String)
        Get_Excel(value)
        ExcelToSQL.table = ExcelToSQL.tb_SQL_Table.Text
        TableGenerator() ' Will Generate Table if does not exist
        ColumnGenerator() ' will Generate Columns if does not exist
        ExcelToSQL.TestResults.Text = "Dropping Information into SQL"
        While worksheet.Cells(row_Num, column_Num).Value <> Nothing
            While column_Num <= column_Count
                content(column_Num - 1) = worksheet.Cells(row_Num, column_Num).Value
                content(column_Num - 1) = Replace(content(column_Num - 1), "'", "''")
                column_Num += 1
                'goes through the data columns and copies the data into an Array
            End While
            total_Records += 1
            SQLDump()
            column_Num = start_Column
            row_Num += 1
            'Resets default columns and rows to access next Excel Sheet
        End While
        row_Num = start_Row
        column_Num = start_Column
        If My.Computer.FileSystem.FileExists("C:\Users\user\Documents\Excel_Log.txt") = False Then
            File.Create("C:\Users\user\Documents\Excel_Log.txt")
            ' If the Log file is not already created, will create file
        Else
        End If
        succeeded_Records = (total_Records - failed_Records)
        My.Computer.FileSystem.WriteAllText("C:\Users\user\Documents\Excel_Log.txt", DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") & " Excel Spreadsheet " & workbook.Name & " has finished aggregating into " & table & " Failed: " & failed_Records.ToString & " Succeeded: " & succeeded_Records.ToString & " Total Tried: " & total_Records.ToString & Environment.NewLine & "======================================================================================" & Environment.NewLine, True)
        workbook.Close()
    End Sub

    Public Sub TableGenerator()
        table = Replace(table, " ", "_")
        table = table.ToLower
        Dim chkString As String = "SELECT TABLE_NAME FROM information_schema.tables WHERE TABLE_NAME = '@TABLE'"
        cmd.CommandText = chkString
        cmd.Parameters.AddWithValue("@TABLE", table)
        Using reader As OleDbDataReader = cmd.ExecuteReader()
            If reader.HasRows Then
                'Table Already Exists
            Else
                'Table Does Not Exist
                ExcelToSQL.TestResults.Text = "Creating Table"
                Dim cmdString As String = "CREATE TABLE "
                cmdString += table & " ("
                While column_Count < col_Max
                    If worksheet.Cells(row_Num, column_Num).Value <> Nothing Then
                        column_Name(column_Count) = worksheet.Cells(row_Num, column_Num).Value
                        column_Name(column_Count) = Replace(column_Name(column_Count), " ", "_") 'convert spaces to _ for SQL column name
                        column_Name(column_Count) = Replace(column_Name(column_Count), "*", "")
                        column_Name(column_Count) = Replace(column_Name(column_Count), "-", "_")
                        column_Name(column_Count) = Replace(column_Name(column_Count), "&", "")
                        column_Name(column_Count) = Replace(column_Name(column_Count), "†", "+")
                        column_Name(column_Count) = Replace(column_Name(column_Count), "–", "_")
                        column_Name(column_Count) = Replace(column_Name(column_Count), "-", "_")
                        column_Count += 1
                        column_Num += 1
                    Else
                        column_Count += 1
                    End If
                End While
                Dim i As Integer = 0
                While i < (column_Count - 1)
                    cmdString += "@COLUMN" & i.ToString & " varchar(255), "
                    cmd.CommandText = cmdString
                    cmd.Parameters.AddWithValue("@COLUMN" & i.ToString, column_Name(i))
                    i += 1
                End While
                cmdString += "@COLUMN" & i.ToString & " varchar(255));"
                reader.Close()
                cmd.CommandText = cmdString
                cmd.Parameters.AddWithValue("@COLUMN" & i.ToString, column_Name(i))
                cmd.ExecuteNonQuery()
            End If
        End Using
    End Sub

    Public Sub ColumnGenerator()
        'Check to make sure Headers from Excel are in Table, Otherwise Create column value
        ' Dim chkString As String = "SELECT COLUMN_NAME FROM information_schema.columns WHERE TABLE_NAME = '" & table & "'"
        column_Count = 0
        While column_Count < col_Max
            If worksheet.Cells(ExcelToSQL.row_Num, column_Num).Value <> Nothing Then
                column_Name(column_Count) = worksheet.Cells(row_Num, column_Num).Value
                column_Name(column_Count) = Replace(column_Name(column_Count), " ", "_") 'convert spaces to _ for SQL column name
                column_Name(column_Count) = Replace(column_Name(column_Count), "*", "")
                column_Name(column_Count) = Replace(column_Name(column_Count), "&", "")
                column_Name(column_Count) = Replace(column_Name(column_Count), "†", "+")
                column_Name(column_Count) = Replace(column_Name(column_Count), "–", "_")
                column_Name(column_Count) = Replace(column_Name(column_Count), "-", "_")
                column_Count += 1
                column_Num += 1
            Else
                column_Count += 1
            End If
        End While
        ExcelToSQL.TestResults.Text = "Creating Fields"
        For Each field_Name In column_Name
            Dim chkString As String = "SELECT COLUMN_NAME FROM information_schema.columns WHERE TABLE_NAME = '@TABLE' AND COLUMN_NAME = '@COLUMN'"
            Dim chkcmd As OleDbCommand = New OleDbCommand(chkString, ExcelToSQL.con)
            chkcmd.Parameters.AddWithValue("@TABLE", table)
            chkcmd.Parameters.AddWithValue("@COLUMN", field_Name)
            Using reader As OleDbDataReader = chkcmd.ExecuteReader()
                If reader.HasRows Then
                    'Field Exists
                Else
                    'Field Does Not Exist
                    Try
                        Dim cmdString As String = "ALTER TABLE @TABLE ADD @FIELD varchar(255);"
                        cmd.CommandText = cmdString
                        cmd.Parameters.AddWithValue("@TABLE", table)
                        cmd.Parameters.AddWithValue("@FIELD", field_Name)
                        cmd.ExecuteNonQuery()
                    Catch ex As Exception

                    Finally
                    End Try
                End If
                reader.Close()
            End Using
        Next
        column_Num = start_Column
        row_Num += 1
    End Sub

    Private Sub SQLDump()
        Dim chkString As String = "SELECT * FROM @TABLE WHERE @COLUMN='@CONTENT'"
        Dim cmd1 As OleDbCommand = New OleDbCommand(chkString, ExcelToSQL.con)
        cmd1.Parameters.AddWithValue("@TABLE", table)
        cmd1.Parameters.AddWithValue("@COLUMN", column_Name(0))
        cmd1.Parameters.AddWithValue("@CONTENT", content(0))
        Using reader As OleDbDataReader = cmd1.ExecuteReader()
            If reader.HasRows Then
                'Record Already Exists
                Try
                    Dim x As Integer = 0
                    Dim dbProvider = "PROVIDER=SQLOLEDB;"
                    Dim dbSource = "Data Source= " & database & ";"
                    Dim sql = "UPDATE @TABLE SET("
                    Using con2 = New OleDb.OleDbConnection(dbProvider & dbSource)
                        Using cmd2 = New OleDb.OleDbCommand(sql, ExcelToSQL.con)
                            cmd2.Parameters.AddWithValue("@TABLE", table)
                            While x < (column_Count - 1)
                                sql += "@Column" & x.ToString & "='@Content" & x.ToString & "', "
                                cmd2.Parameters.AddWithValue("@Column" & x.ToString, column_Name(x))
                                cmd2.Parameters.AddWithValue("@Content" & x.ToString, content(x))
                                x += 1
                            End While
                            sql += "@Column" & x.ToString & "='@Content" & x.ToString & "' "
                            cmd2.Parameters.AddWithValue("@Column" & x.ToString, column_Name(x))
                            cmd2.Parameters.AddWithValue("@Content" & x.ToString, content(x))
                            sql += "WHERE @Column0 = @Content0"
                            cmd2.Parameters.AddWithValue("@Column0", column_Name(0))
                            cmd2.Parameters.AddWithValue("@Content0", content(0))
                            x = 0
                            cmd.ExecuteNonQuery()
                        End Using
                    End Using
                Catch ex As Exception
                    If My.Computer.FileSystem.FileExists("C:\Users\user\Documents\Excel_Log.txt") = False Then
                        File.Create("C:\Users\user\Documents\Excel_Log.txt")
                        ' If the Log file is not already created, will create file
                    Else
                    End If
                    failed_Records += 1
                    My.Computer.FileSystem.WriteAllText("C:\Users\user\Documents\Excel_Log.txt", DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") & " Excel Spreadsheet " & workbook.Name & " has Error Aggregating into " & table & " " & ex.Message & Environment.NewLine & "====================================================================================" & Environment.NewLine, True)
                Finally
                End Try
            Else
                'Record does not Exist
                Try
                    Dim x As Integer = 0
                    Dim dbProvider = "PROVIDER=SQLOLEDB;"
                    Dim dbSource = "Data Source= " & database & ";"
                    Dim sql = "INSERT INTO @TABLE ("
                    Using con2 = New OleDb.OleDbConnection(dbProvider & dbSource)
                        Using cmd2 = New OleDb.OleDbCommand(sql, ExcelToSQL.con)
                            cmd2.Parameters.AddWithValue("@TABLE", table)
                            While x < (column_Count - 1)
                                sql += "@Column" & x.ToString & ", "
                                cmd2.Parameters.AddWithValue("@Column" & x.ToString, column_Name(x))
                                x += 1
                            End While
                            sql += "@Column" & x.ToString & ") VALUES ("
                            cmd2.Parameters.AddWithValue("@Column" & x.ToString, column_Name(x))
                            x = 0
                            While x < (column_Count - 1)
                                sql += "@Content" & x.ToString & ", "
                                cmd2.Parameters.AddWithValue("@Content" & x.ToString, content(x))
                                x += 1
                            End While
                            x = 0
                            cmd.ExecuteNonQuery()
                        End Using
                    End Using
                Catch ex As Exception
                    If My.Computer.FileSystem.FileExists("C:\Users\user\Documents\Excel_Log.txt") = False Then
                        File.Create("C:\Users\user\Documents\Excel_Log.txt")
                        ' If the Log file is not already created, will create file
                    Else
                    End If
                    failed_Records += 1
                    My.Computer.FileSystem.WriteAllText("C:\Users\user\Documents\Excel_Log.txt", DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") & " Excel Spreadsheet " & workbook.Name & " has Error Aggregating into " & table & " " & ex.Message & Environment.NewLine & "====================================================================================" & Environment.NewLine, True)
                Finally

                End Try
            End If
            reader.Close()
        End Using
    End Sub
End Class
