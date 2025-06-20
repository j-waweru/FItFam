Imports System.Data.OleDb
Imports System.Diagnostics.Eventing.Reader

Public Class Registration
    Dim mycon As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\USER\Documents\KCAU\2.2 sss 3.1 gss\project\Coding\FitFam\Backend.accdb")
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click, Label2.Click

    End Sub

    Private Sub Registration_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        mycon.Open()

        Dim query As String = "INSERT INTO UserTable (FullName,UserName,Pword,Gender,Email,FitnessGoals) VALUES (?, ?, ?, ?, ?, ?)"
        Dim mycmd As New OleDbCommand(query, mycon)

        mycmd.Parameters.AddWithValue("@FullName", TextBox1.Text)
        mycmd.Parameters.AddWithValue("@UserName", TextBox3.Text)
        mycmd.Parameters.AddWithValue("@Pword", TextBox2.Text)
        mycmd.Parameters.AddWithValue("@Gender", ComboBox1.SelectedItem.ToString)
        mycmd.Parameters.AddWithValue("@Email", TextBox5.Text)
        mycmd.Parameters.AddWithValue("@FitnessGoals", ComboBox2.SelectedItem.ToString)

        Try
            mycmd.ExecuteNonQuery()

            TextBox1.Clear()
            TextBox3.Clear()
            TextBox2.Clear()
            TextBox5.Clear()
            ComboBox1.Items.Clear()
            ComboBox2.Items.Clear()
            mycon.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Create an instance of the Bookingsytem to allow the user to book
        Me.Close()
        Login.Show()


    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class