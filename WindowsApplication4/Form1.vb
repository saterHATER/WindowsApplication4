Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.OleDb

Public Class Form1
    Inherits System.Windows.Forms.Form
    Dim OleConn As New OleDb.OleDbConnection
    Dim OleReader As OleDb.OleDbDataReader
    Dim OleReader1 As OleDb.OleDbDataReader
    Dim sqlConn As SqlConnection
    Dim Cmd As New SqlCommand
    Dim myReader As SqlDataReader
    Dim results As String
    Dim textbox1 As New TextBox
    Dim oriGin As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    Dim secSchtuff As String = ";Persist Security Info=True;Jet OLEDB:Database Password=time2634"
    'Dim bawsQuery = "SELECT MSysObjects.Name AS table_name FROM MSysObjects WHERE (((Left([Name],1))<>""~"") AND ((Left([Name],4))<>""MSys"") AND ((MSysObjects.Type) In (1,4,6))) order by MSysObjects.Name"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Create an sql connecction object
        sqlConn = New SqlConnection("Data Source=" & Environment.MachineName & "; Initial Catalog=master; Integrated Security=True;")

        'Grab data from SQL database
        Dim myCmd As SqlDataAdapter = New SqlDataAdapter("EXEC sp_databases", sqlConn)
        Dim myData As New DataSet("tpDataSQL")

        'Makin' a string array
        Dim nameList As New List(Of String)

        sqlConn.Open()

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
            textbox1.Text = dialog.FileName
            OleConn.ConnectionString = oriGin & textbox1.Text & secSchtuff
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
        konnekt()
    End Sub

    Function konnekt()
        Try
            sqlConn.Open()
            Console.WriteLine(sqlConn.State)
            OleConn.Open()
            Console.WriteLine(OleConn.State)

            'GET TABLE NAMES
            'Dim allTableQuery As String = "SELECT * FROM INFORMATION_SCHEMA.TABLES"
            'Dim allTableCmd As SqlDataAdapter = New SqlDataAdapter(allTableQuery, sqlConn)
            'Dim allTableSet As New DataSet()
            'allTableCmd.Fill(allTableSet)
            'For Each table As DataTable In allTableSet.Tables
            '    Dim currentTable1 As String = table.Namespace
            'Next
            Dim sr As New System.IO.StreamReader("tableNames.txt")
            While sr.Read()
                Dim currentTable As String = "T" & sr.ReadLine()
                'Console.WriteLine(currentTable)
                'ITERATE OVER TABLE
                'Dim currentTable = "tblJob"
                Dim OleCmd = New OleDb.OleDbCommand("SELECT * FROM " & currentTable, OleConn) '************good
                OleReader = OleCmd.ExecuteReader()

                'GO THROUGH WHOLE ROW
                While OleReader.Read()
                    Dim datum As String = Nothing
                    Dim fieldNames As String = Nothing

                    'GO THROUGH EACH COLUMN
                    For columns = 0 To OleReader.FieldCount - 1
                        datum = datum & ", '" & OleReader.Item(columns) & "'"
                        fieldNames = fieldNames & ", " & OleReader.GetName(columns)
                        'Console.WriteLine(datum)
                    Next
                    datum = datum.Substring(2, (datum.Length - 2))
                    fieldNames = fieldNames.Substring(2, (fieldNames.Length - 2))
                    Cmd.CommandText = "INSERT INTO " & currentTable & " (" & fieldNames & ")" & " VALUES (" & datum & ")"
                    'Console.WriteLine(Cmd.CommandText)
                    Cmd.Connection = sqlConn
                    Cmd.ExecuteNonQuery()
                    Console.ReadLine()
                End While
            End While
        Catch ex As Exception
            MsgBox("oops: " & vbCrLf & ex.Message)
        End Try
        OleConn.Close()
        sqlConn.Close()
        Return Nothing
    End Function

End Class