Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.OleDb
Imports System.IO

Public Class Form1
    Inherits System.Windows.Forms.Form
    Dim OleConn As New OleDb.OleDbConnection    'We hate global variables but we need to persist over the whole app life
    Dim sqlConn As SqlConnection                'Must leave these connection variables as global persistant variables
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
            MsgBox("Bad SQL DB: " & vbCrLf & ex.Message & vbCrLf & sqlConn.ConnectionString)
        End Try
        GoOverDB(True)
    End Sub

    Function GoOverDB(ByVal readOrRemove As Boolean)
        Try
            sqlConn.Open()
            OleConn.Open()
            Dim FileContents As String = File.ReadAllText(Application.StartupPath & "\tableNames.txt")
            Dim tableArray() As String = FileContents.Split(vbCrLf)
            'ITERATE OVER TABLE
            For I = 0 To tableArray.GetUpperBound(0)
                Dim currentTable As String = tableArray(I)
                Console.WriteLine("table: " & currentTable)
                If readOrRemove Then
                    GoOverTable(currentTable)
                Else
                    clearTable(currentTable)
                End If
            Next
        Catch ex As Exception
            MsgBox("darn, the connection broke: " & vbCrLf & ex.Message)
        End Try
        OleConn.Close()
        sqlConn.Close()
        Return Nothing
    End Function

    Function GoOverTable(ByVal tableName As String)
        Try
            Dim OleCmd = New OleDb.OleDbCommand("SELECT * FROM " & tableName, OleConn) 'grab all data from our access DB
            Dim OleReader As OleDbDataReader = OleCmd.ExecuteReader() 'grab all data from our access DB
            Dim rowCount As Int16 = 1
            While OleReader.Read()
                goOverRow(OleReader, tableName, rowCount)
                rowCount += 1
            End While
        Catch es As Exception
            MsgBox("Arrg! reading " & tableName & " had a snag: " & vbCrLf & es.Message)
        End Try
        Return Nothing
    End Function

    Function goOverRow(ByVal oleReader As OleDbDataReader, ByVal tableName As String, ByVal rowNum As Int16)
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
            MsgBox("Aw, Snaps! We couldn't write to SQL!: " & vbCrLf & el.Message & vbCrLf & dataRow & "continue?" &, 1)
            'If response = MsgBoxResult.Yes Then
            '    vbAbort()
            'End If
        End Try
        Return Nothing
    End Function

    Function clearTable(ByVal tblName As String)
        Try
            Dim sqlCmd As New SqlCommand
            sqlCmd.CommandText = "DELETE FROM " & tblName
            sqlCmd.Connection = sqlConn
            sqlCmd.ExecuteNonQuery()
        Catch en As Exception
            MsgBox("Couldn't Even Delete the table: " & vbCrLf & en.Message)
        End Try
        Return Nothing
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim curItem As String = ListBox1.SelectedItem.ToString()
            sqlConn.ConnectionString = "Data Source=" & Environment.MachineName & "; Initial Catalog=" & ListBox1.SelectedItem.ToString() & "; Integrated Security=True;"
            Console.WriteLine(sqlConn.ConnectionString)
        Catch ex As Exception
            MsgBox("oops: " & vbCrLf & ex.Message)
        End Try
        GoOverDB(False)
    End Sub

End Class