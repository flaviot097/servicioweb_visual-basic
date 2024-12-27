Imports System.Web
Imports System.Data.SqlClient
Public Class Conexion
    Implements IHttpModule

    Private WithEvents _context As HttpApplication

    Dim cadena_de_coxion As String = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=App_gestor_caja;Data Source=DESKTOP-SEVLBL8"

    Dim conexion As New SqlConnection(cadena_de_coxion)

    Public Sub CrearConexion()
        Try
            conexion.Open()
            Dim consulta As String = "SELECT * FROM Users"
            Dim cmd As New SqlCommand(consulta, conexion)
            Dim lector As SqlDataReader = cmd.ExecuteReader()

            While lector.Read()
                Console.WriteLine(lector("Id") + " " + lector("Username") + " " + lector("password"))
            End While
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    '''  Deberá configurar este módulo en el archivo web.config de su
    '''  web y registrarlo en IIS para poder usarlo. Para obtener más información
    '''  consulte el vínculo siguiente: https://go.microsoft.com/?linkid=8101007
    ''' </summary>
#Region "Miembros de IHttpModule"

    Public Sub Dispose() Implements IHttpModule.Dispose

        ' Ponga aquí el código de limpieza

    End Sub

    Public Sub Init(ByVal context As HttpApplication) Implements IHttpModule.Init
        _context = context
    End Sub

#End Region

    Public Sub OnLogRequest(ByVal source As Object, ByVal e As EventArgs) Handles _context.LogRequest

        ' Controla el evento LogRequest para proporcionar una implementación de 
        ' registro personalizado para él

    End Sub
End Class
