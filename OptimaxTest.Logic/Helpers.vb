Public Class Helpers
    Public Shared Function ConvertAppUserToUser(
            appUserID As Integer,
            username As String,
            passwordHash As Byte(),
            firstName As String,
            surname As String,
            userRoleID As Integer,
            isActive As Boolean,
            dateTimeCreated As DateTime,
            dateTimeDeactivated? As DateTime
        ) As Object

        Dim user As New Object
        With user
            .UserID = appUserID
            .UserName = username
            .Password = ConvertFromBinaryToString(passwordHash)
            .DisplayName = $"{firstName} {surname}"
            .UserRoleID = userRoleID
            .Active = isActive
            .Created = dateTimeCreated
        End With

        Return user
    End Function

    Public Shared Function ConvertFromBinaryToString(hash As Byte())
        Return Text.Encoding.Unicode.GetString(hash)
    End Function

    Public Shared Function ConvertStringToBinary(str As String)
        Return Text.Encoding.Unicode.GetBytes(str)
    End Function
End Class
