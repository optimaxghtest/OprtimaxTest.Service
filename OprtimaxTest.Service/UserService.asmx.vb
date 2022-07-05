Imports System.ComponentModel
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports OptimaxTest.Logic

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class UserService
    Inherits System.Web.Services.WebService
    Private mLogic As UserServiceLogic

    Public Sub New()

        mLogic = New UserServiceLogic()

    End Sub


    <WebMethod(Description:="Get app user by passing the app user's ID",
            MessageName:="GetAppUserByUserID")>
    Public Function GetUserByUserID(appUserID As Integer) As DataSet

        mLogic.UserHostAddress = Context.Request.UserHostAddress()
        Return mLogic.GetAppUserByAppUserID(appUserID)

    End Function

    <WebMethod(Description:="Get a list of all app users",
            MessageName:="GetAllAppUsers")>
    Public Function GetAllAppUsers() As DataSet

        mLogic.UserHostAddress = Context.Request.UserHostAddress()
        Return mLogic.GetAllAppUsers()

    End Function



    <WebMethod(Description:="Check if user hold permission by passing the app users ID and permission ID",
            MessageName:="CheckAppUserPermissionByPermissionID")>
    Public Function CheckAppUserPermissionByPermissionID(appUserID As Integer, permissionID As Integer) As Boolean

        mLogic.UserHostAddress = Context.Request.UserHostAddress()
        Return mLogic.CheckAppUserPermissionByPermissionID(appUserID, permissionID)

    End Function

    <WebMethod(Description:="Insert new user into the AppUser table",
            MessageName:="InsertNewUser")>
    Public Sub InsertNewUser(
                            username As String,
                            password As String,
                            firstName As String,
                            surname As String,
                            userRoleID As Integer,
                            isActive As Boolean
                        )

        mLogic.UserHostAddress = Context.Request.UserHostAddress()
        mLogic.InsertNewUser(username, password, firstName, surname, userRoleID, isActive)
    End Sub

End Class