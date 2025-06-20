Imports System.Data.OleDb

Public Class UserDashboard
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Create an instance of the booking system
        Me.Close()
        Booking.Show()


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub UserDashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'get the userid from username
        Dim username As String = Login.TextBox3.Text

        Using myCon As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\USER\Documents\KCAU\2.2 sss 3.1 gss\project\Coding\FitFam\Backend.accdb")
            myCon.Open()
            Dim userID As Integer
            Dim eqID As Integer

            ' Query to get UserID from UserTable
            Dim query As String = "SELECT UserID FROM UserTable WHERE Username = @Username"
            Using command As New OleDbCommand(query, myCon)
                command.Parameters.AddWithValue("@Username", username)
                Dim result As Object = command.ExecuteScalar()
                If result IsNot DBNull.Value Then
                    userID = Convert.ToInt32(result)
                End If
            End Using

            ' Query to get EquipmentID for the UserID booked
            Dim query2 As String = "SELECT EquipmentID FROM BookingTable WHERE UserID = @UserID"
            Using command2 As New OleDbCommand(query2, myCon)
                command2.Parameters.AddWithValue("@UserID", userID)
                Dim result2 As Object = command2.ExecuteScalar()
                If result2 IsNot DBNull.Value Then
                    eqID = Convert.ToInt32(result2)
                End If
            End Using

            ' Query to get EquipmentName from EquipmentTable for the EquipmentID
            Dim query3 As String = "SELECT EquipmentName FROM EquipmentTable WHERE EquipmentID = @eqID"
            Using command3 As New OleDbCommand(query3, myCon)
                command3.Parameters.AddWithValue("@eqID", eqID)
                Dim result3 As OleDbDataReader = command3.ExecuteReader()
                If result3.Read() Then
                    ListBox1.Items.Add(result3("EquipmentName").ToString())
                Else
                    ListBox1.Items.Add("You have no bookings")
                End If
            End Using
        End Using
    End Sub


End Class