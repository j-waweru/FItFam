Imports System.Data.OleDb

Public Class Login
    Dim myCon As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\USER\Documents\KCAU\2.2 sss 3.1 gss\project\Coding\FitFam\Backend.accdb")
    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Registration.Show()
        Me.Hide()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        myCon.Open()

        Dim mycmd As New OleDbCommand("SELECT UserName,Pword  FROM UserTable WHERE UserName= '" & TextBox3.Text & "' And  Pword = '" & TextBox1.Text & "'", myCon)
        Dim myread As OleDbDataReader = mycmd.ExecuteReader
        If myread.Read Then
            MsgBox("Successful login")
            TextBox1.Clear()

            ' Create an instance of the Bookingsytem to allow the user to book
            Booking.Show()

        Else
            MsgBox("User Name or Password is Incorrect")
        End If
        myCon.Close()

    End Sub

    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Label1_Click_1(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub
End Class
