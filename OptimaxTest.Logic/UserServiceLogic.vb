Imports OptimaxTest.Data
Imports System

Public Class UserServiceLogic

    'TODO: ADD DB CONNECTION
    Private Const cachePeriod As Integer = 60
    Private mUserHostAddress As String

    Public Property UserHostAddress As String

        Get
            Return mUserHostAddress
        End Get

        Set(value As String)
            mUserHostAddress = value
        End Set

    End Property

    Public Function GetAppUserByAppUserID(appUserID As Integer) As DataSet
        Dim sqlCmd As SqlClient.SqlCommand
        Dim param As SqlClient.SqlParameter
        Dim userData As DataSet

        userData = Nothing

        sqlCmd = New SqlClient.SqlCommand
        sqlCmd.CommandType = CommandType.StoredProcedure
        sqlCmd.CommandText = "GetAppUserByAppUserID"

        param = sqlCmd.Parameters.AddWithValue("@AppUserID", appUserID)

        Try
            userData = GetDataSet(sqlCmd)
        Catch ex As Exception
            'imagine error logging functionality here
        End Try

        Return userData

    End Function

    Public Function GetAllAppUsers() As DataSet
        Dim sqlCmd As SqlClient.SqlCommand
        Dim userData As DataSet

        userData = Nothing

        sqlCmd = New SqlClient.SqlCommand
        sqlCmd.CommandType = CommandType.StoredProcedure
        sqlCmd.CommandText = "GetAllAppUsers"


        Try
            userData = GetDataSet(sqlCmd)
        Catch ex As Exception
            'imagine error logging functionality here
        End Try

        Return userData
    End Function

    Public Function CheckAppUserPermissionByPermissionID(appUserID As Integer, permissionID As Integer) As Boolean
        Dim sqlCmd As SqlClient.SqlCommand
        Dim param As SqlClient.SqlParameter
        Dim permissionData As DataSet
        Dim hasPermission As Boolean = Nothing

        permissionData = Nothing

        sqlCmd = New SqlClient.SqlCommand
        sqlCmd.CommandType = CommandType.StoredProcedure
        sqlCmd.CommandText = "CheckAppUserPermissionByPermissionID"

        param = sqlCmd.Parameters.Add("@HasPermission", SqlDbType.Bit)
        param.Value = DBNull.Value
        param.Direction = ParameterDirection.Output

        param = sqlCmd.Parameters.AddWithValue("@AppUserID", appUserID)
        param = sqlCmd.Parameters.AddWithValue("@PermissionID", permissionID)


        Try
            ExecuteSQL(sqlCmd)

            hasPermission = Convert.ToBoolean(sqlCmd.Parameters("@HasPermission").Value())
        Catch ex As Exception
            'imagine error logging functionality here
        End Try


        Return hasPermission

    End Function

    Public Sub InsertNewUser(
                            username As String,
                            password As String,
                            firstName As String,
                            surname As String,
                            userRoleID As Integer,
                            isActive As Boolean
                        )
        Try
            InsertUser(
               username,
               password,
               firstName,
               surname,
               userRoleID,
               isActive
            )
        Catch ex As Exception
            'imagine error logging here
        End Try
    End Sub

    Public Sub InsertUser(
                        username As String,
                        password As String,
                        firstName As String,
                        surname As String,
                        userRoleID As Integer,
                        isActive As Boolean
                        )
        Dim sqlCmd As SqlClient.SqlCommand
        Dim param As SqlClient.SqlParameter
        Dim returnID As Integer

        returnID = 0

        sqlCmd = New SqlClient.SqlCommand
        sqlCmd.CommandType = CommandType.StoredProcedure
        sqlCmd.CommandText = "InsertUser"

        param = sqlCmd.Parameters.AddWithValue("@UserName", username.ToUpper())
        param = sqlCmd.Parameters.AddWithValue("@Password", Helpers.ConvertStringToBinary(password))
        param = sqlCmd.Parameters.AddWithValue("@FirstName", firstName)
        param = sqlCmd.Parameters.AddWithValue("@Surname", surname)
        param = sqlCmd.Parameters.AddWithValue("@UserRoleID", userRoleID)
        param = sqlCmd.Parameters.AddWithValue("@IsActive", isActive)

        Try
            ExecuteSQL(sqlCmd)

        Catch ex As Exception
            'imagine error logging here
        End Try
    End Sub
End Class
