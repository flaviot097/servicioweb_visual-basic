Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.IO
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml

' Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://servicioweb.com/miServicio")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class servicioweb
    Inherits System.Web.Services.WebService

    Private cadena_de_coxion As String = "server=DESKTOP-SEVLBL8; database=App_gestor_caja;integrated security=true;"

    Private conexion As New SqlConnection(cadena_de_coxion)

    <WebMethod()>
    Public Function HelloWorld() As String
        Return "Hola a todos"
    End Function

    <WebMethod()>
    Public Function CrearConsulta() As XmlDocument

        'Dim resultado As String = ""

        ' Creo un documento xml vacio
        Dim documentoXml As XmlDocument = New XmlDocument()

        ' Creo el elemento usuario y posteriormente lo Inserto en el documento
        Dim raiz As XmlElement = documentoXml.CreateElement("Usuarios")
        documentoXml.AppendChild(raiz)
        Dim consultaSQL As String = "SELECT * FROM Users"

        Try
            conexion.Open()
            Dim comando As New SqlCommand(consultaSQL, conexion)
            Dim lector As SqlDataReader = comando.ExecuteReader()

            While lector.Read()

                ' creo el nodo Usuario
                Dim Usuario As XmlElement = documentoXml.CreateElement("Usuario")

                ' Agrego elementos al Nodo Usuario
                Dim id As XmlElement = documentoXml.CreateElement("ID")
                ' le doy un valor a este elemento
                id.InnerText = lector("id")
                ' Añedo el elemento al nodo usuario
                Usuario.AppendChild(id)

                Dim nombreUsuario As XmlElement = documentoXml.CreateElement("Nombre_de_usuario")
                nombreUsuario.InnerText = lector("Username")
                Usuario.AppendChild(nombreUsuario)

                Dim contrasena As XmlElement = documentoXml.CreateElement("Contraseña")
                contrasena.InnerText = lector("password")
                Usuario.AppendChild(contrasena)

                ' una vez terminado y habiendole agregados elementos al nodo agrego el nodo usuario a la raiz del xml
                raiz.AppendChild(Usuario)
                ' resultado &= lector("Username").ToString() + "; " & lector("id") & "; " & +lector("password").ToString() & vbCrLf
            End While
            lector.Close()

        Catch ex As Exception
            ' Devuelve el error
            Dim errorNodo As XmlElement = documentoXml.CreateElement("Error")
            errorNodo.InnerText = ex.Message
            raiz.AppendChild(errorNodo)
        Finally
            If conexion.State = ConnectionState.Open Then
                conexion.Close()
            End If
        End Try

        Return documentoXml
    End Function

    <WebMethod()>
    Public Function SeleccionarUsuario(id As Integer) As XmlDocument

        Dim valor = id.ToString()
        Dim consultaSLQ As String = "SELECT id,Username,password FROM Users WHERE id = @id"
        Dim DocumentoXML As XmlDocument = New XmlDocument()

        Dim raiz As XmlElement = DocumentoXML.CreateElement("Usuario")
        DocumentoXML.AppendChild(raiz)

        Try
            conexion.Open()
            Dim cmd As SqlCommand = New SqlCommand(consultaSLQ, conexion)
            cmd.Parameters.AddWithValue("@id", valor)
            Dim lectora As SqlDataReader = cmd.ExecuteReader()

            ' Si hay usuarios haceme esto:
            If lectora.HasRows Then
                While lectora.Read()
                    Dim UsuarioXML As XmlElement = DocumentoXML.CreateElement($"Usuario {lectora("Username")}")

                    Dim idUser As XmlElement = DocumentoXML.CreateElement("id")
                    idUser.InnerText = lectora("Id").ToString()
                    UsuarioXML.AppendChild(idUser)

                    Dim nombre As XmlElement = DocumentoXML.CreateElement("Nombre_de_usuario")
                    nombre.InnerText = lectora("Username").ToString()
                    UsuarioXML.AppendChild(nombre)

                    Dim password As XmlElement = DocumentoXML.CreateElement("Contraseña")
                    password.InnerText = lectora("password").ToString()
                    UsuarioXML.AppendChild(password)

                    raiz.AppendChild(UsuarioXML)
                End While

                'Sino haceme esto:
            Else
                Dim mensaje As XmlElement = DocumentoXML.CreateElement("Mensaje")
                mensaje.InnerText = "Usuario no encontrado"
                raiz.AppendChild(mensaje)
            End If
            lectora.Close()

        Catch ex As Exception
            Dim errorNodo As XmlElement = DocumentoXML.CreateElement("Error")
            errorNodo.InnerText = ex.Message
            raiz.AppendChild(errorNodo)
        Finally
            If conexion.State = ConnectionState.Open Then
                conexion.Close()
            End If
        End Try

        Return DocumentoXML

    End Function

    <WebMethod()>
    Public Function crearUsuario(nombre As String, contrasena As String) As XmlDocument

        Dim documentoRespuesta As New XmlDocument()
        Dim raiz As XmlElement = documentoRespuesta.CreateElement("Resultado_de_insercion")
        Dim respuesta As String

        Dim cosultaSql As String = "INSERT INTO Users (Username ,password) VALUES (@username, @password)"

        Try
            conexion.Open()
            Dim cmd As SqlCommand = New SqlCommand(cosultaSql, conexion)
            cmd.Parameters.AddWithValue("@username", nombre)
            cmd.Parameters.AddWithValue("@password", contrasena)

            Dim filasAfectadas As Integer = cmd.ExecuteNonQuery()
            If filasAfectadas > 0 Then
                respuesta = "Se a creado correctamentre el usuario"
            Else
                respuesta = "No se pudo crear el usuario"
            End If
        Catch ex As Exception
            respuesta = "Error:" & ex.Message
        End Try

        raiz.InnerText = respuesta
        documentoRespuesta.AppendChild(raiz)

        Return documentoRespuesta

    End Function

    <WebMethod()>
    Public Function EliminarUsuario(id As Integer) As XmlDocument
        Dim consulat As String = $"DELETE FROM Users WHERE id={id}"
        Dim documentoXML As XmlDocument = New XmlDocument()
        Dim raiz As XmlElement = documentoXML.CreateElement("Resultado_de_eliminacion")
        Dim respuesta As String
        Try
            conexion.Open()
            Dim cmd As New SqlCommand(consulat, conexion)
            Dim filasAfectadas As Integer = cmd.ExecuteNonQuery()

            If filasAfectadas > 0 Then
                respuesta = "El usuario se elimino con exito"
            Else
                respuesta = "No se encontro un usuario con ese id ,intente nuevamente"
            End If

        Catch ex As Exception
            respuesta = ex.Message()
        End Try
        raiz.InnerText = respuesta
        documentoXML.AppendChild(raiz)

        Return documentoXML
    End Function

    <WebMethod()>
    Public Function ActualizarUsuario(id As Integer, nombre As String, contrasena As String) As XmlDocument
        Dim resultado As String
        Dim documentoXML As XmlDocument = New XmlDocument()
        Dim raiz As XmlElement = documentoXML.CreateElement("Resultado_de_actualiazar")
        'Dim consulta As String = $"UPDATE Users SET Username = {nombre} , password= {contrasena} WHERE id={id}"
        Dim consulta As String = $"UPDATE Users SET Username =@username ,password=@password WHERE id=@id"


        Try
            conexion.Open()
            Dim cmd As SqlCommand = New SqlCommand(consulta, conexion)
            cmd.Parameters.AddWithValue("@username", nombre)
            cmd.Parameters.AddWithValue("@password", contrasena)
            cmd.Parameters.AddWithValue("@id", id)

            Dim filasAfectadas As Integer = cmd.ExecuteNonQuery()
            If filasAfectadas > 0 Then
                resultado = "Se actualizo el usuario con exito"
            Else
                resultado = "Error al actializar usuario"
            End If
        Catch ex As Exception
            resultado = "Error " & ex.Message
        Finally
            If conexion.State = ConnectionState.Open Then
                conexion.Close()
            End If
        End Try

        raiz.InnerText = resultado
        documentoXML.AppendChild(raiz)

        Return documentoXML
    End Function
End Class