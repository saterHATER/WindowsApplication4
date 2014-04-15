Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.OleDb

Public Class Form1
    Inherits System.Windows.Forms.Form
    Dim OleConn As New OleDb.OleDbConnection    'We hate global variables but we need to persist over the whole app life
    Dim sqlConn As SqlConnection                'must leave these connection variables as global persistant variables
    Dim oriGin As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="                          'ugh, these are fine as well
    Dim secSchtuff As String = ";Persist Security Info=True;Jet OLEDB:Database Password=time2634"   'ugh, these are fine as well

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Create an sql connecction object
        sqlConn = New SqlConnection("Data Source=" & Environment.MachineName & "; Initial Catalog=master; Integrated Security=True;")

        'Grab data from SQL database
        Dim myCmd As SqlDataAdapter = New SqlDataAdapter("EXEC sp_databases", sqlConn)
        Dim myData As New DataSet("tpDataSQL")

        'Makin' a string array
        Dim nameList As New List(Of String)

        sqlConn.Open() ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'READ FROM ACCESS DATABASE
        myCmd.FillSchema(myData, SchemaType.Source, "tblJob")
        myCmd.Fill(myData, "tblJob")

        'MAKE A DATATABLE
        Dim tblJob As DataTable
        tblJob = myData.Tables("tblJob")

        'DISPLAY DATATABLE
        Dim drCurrent As DataRow
        For Each drCurrent In tblJob.Rows
            ListBox1.Items.Add(drCurrent("DATABASE_NAME"))
        Next
        Console.ReadLine()
        sqlConn.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim dialog As New OpenFileDialog()
        dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        dialog.Filter = "mdb files (*.mdb)|*.mdb|All files (*.*)|*.*"
        If DialogResult.OK = dialog.ShowDialog Then
            OleConn.ConnectionString = oriGin & dialog.FileName & secSchtuff
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim curItem As String = ListBox1.SelectedItem.ToString()
            sqlConn.ConnectionString = "Data Source=" & Environment.MachineName & "; Initial Catalog=" & ListBox1.SelectedItem.ToString() & "; Integrated Security=True;"
            Console.WriteLine(sqlConn.ConnectionString)
        Catch ex As Exception
            MsgBox("oops: " & vbCrLf & ex.Message)
        End Try
        DBmuncher()
    End Sub

    Function DBmuncher()
        Try
            sqlConn.Open()
            OleConn.Open()
            Dim sr As New System.IO.StreamReader("tableNames.txt")
            'ITERATE OVER TABLE
            While sr.Read()
                Dim currentTable As String = "T" & sr.ReadLine() 'haven't quite fixed the readline() issue with sr
                everyTable(currentTable)
            End While
        Catch ex As Exception
            MsgBox("Yowchers! We didn't finish!: " & vbCrLf & ex.Message)
        End Try
        OleConn.Close()
        sqlConn.Close()
        Return Nothing
    End Function

    Function everyTable(ByVal tableName As String)
        Try
            Dim OleCmd = New OleDb.OleDbCommand("SELECT * FROM " & tableName, OleConn) 'grab all data from our access DB
            Dim OleReader As OleDbDataReader = OleCmd.ExecuteReader() 'grab all data from our access DB
            Dim rowCount As Int16 = 1
            While OleReader.Read()
                everyRow(OleReader, tableName, rowCount)
                rowCount += 1
            End While
        Catch es As Exception
            MsgBox("Arrg! reading" & tableName & " had a snag: " & vbCrLf & es.Message)
        End Try
        Return Nothing
    End Function

    Function everyRow(ByVal oleReader As OleDbDataReader, ByVal tableName As String, ByVal rowNum As Int16)
        Try
            Dim dataRow As String = Nothing
            Dim fieldNames As String = Nothing
            For columns = 0 To oleReader.FieldCount - 1
                dataRow = dataRow & ", '" & oleReader.Item(columns) & "'"
                fieldNames = fieldNames & ", " & oleReader.GetName(columns)
            Next
            writeToSQL(dataRow, tableName, fieldNames)
        Catch eg As Exception
            MsgBox("Darn! Reading row" & rowNum & " had a snag: " & vbCrLf & eg.Message)
        End Try
        Return Nothing
    End Function

    Function writeToSQL(ByVal dataRow As String, ByVal tableName As String, ByVal fieldNames As String)
        Try
            Dim sqlCmd As New SqlCommand
            dataRow = dataRow.Substring(2, (dataRow.Length - 2))
            fieldNames = fieldNames.Substring(2, (fieldNames.Length - 2))
            sqlCmd.CommandText = "INSERT INTO " & tableName & " (" & fieldNames & ")" & " VALUES (" & dataRow & ")"
            sqlCmd.Connection = sqlConn
            sqlCmd.ExecuteNonQuery()
        Catch el As Exception
            MsgBox("Aw, Snaps! We couldn't write to SQL!: " & vbCrLf & el.Message)
        End Try
        Return Nothing
    End Function

End Class