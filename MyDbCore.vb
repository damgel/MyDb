Imports Microsoft.VisualBasic
Imports System.Configuration
Imports MySql.Data

'/**
'*
'* @author: daMgeL
'* @date:   03-03-2014:2:04PM
'*
'/**

Public Class MyDbCore

    Implements IDisposable
    Private Shared myDataSet As New Data.DataSet
    Private Shared myCnSQL As New MySqlClient.MySqlConnection(GetConnectionStringByName("myConnString"))
    Private Shared myCommandSQL As New MySqlClient.MySqlCommand
    Private Shared Function GetConnectionStringByName(ByVal name As String) As String
        ' asumo que puede ocurrir un error
        Dim returnCnnString As String = Nothing
        ' busco el nombre de la coneccion en el conectionString
        Dim settings As ConnectionStringSettings = ConfigurationManager.ConnectionStrings(name)
        'Si existe, retorno el valor.
        If Not settings Is Nothing Then
            returnCnnString = settings.ConnectionString
        End If
        Return returnCnnString
    End Function
    Private Shared Sub ABM(ByVal statementSQL As String)
        'A=insetar(altas), B=Eliminar(bajas), M=modificar(modificaciones).
        Try
            myCnSQL.Open()
            myCommandSQL = New MySqlClient.MySqlCommand(statementSQL, myCnSQL)
            myCommandSQL.ExecuteNonQuery()
        Catch Ex As Exception
            MsgBox("Error al conectar a la base de datos: " + Ex.Message)
        Finally
            myCnSQL.Close()
        End Try
    End Sub
    Public Shared Sub Insert(statementSQL As String)
        Dim ins As String = "INSERT INTO "
        ABM(ins + statementSQL)
    End Sub
    Public Shared Sub Update(statementSQL As String)
        Dim upd As String = "UPDATE "
        ABM(upd + statementSQL)
    End Sub
    Public Shared Sub Delete(statementSQL As String)
        Dim del As String = "DELETE "
        ABM(del + statementSQL)
    End Sub
    Public Shared Sub DeleteById(statementSQL As String, Where As String)
        Dim del As String = "DELETE "
        Where = "WHERE " + Where
        ABM(del + statementSQL + Where)
    End Sub
    Public Shared Function GetDatosAsDataset( _
                                            ByVal FIELDS As String, _
                                            ByVal FROM As String, _
                                            Optional WHERE As String = "") As Data.DataSet
        'Funcion para buscar datos, permite llenar un dataset y un combo.
        Try
            Dim setWhere As String = WHERE
            If setWhere = "" Then
                'valor asignado en el where
                setWhere = ""
            Else
                setWhere = " WHERE " & WHERE.ToString
            End If
            Dim setFrom As String = " FROM " & FROM.ToString
            myDataSet = New Data.DataSet
            myCnSQL.Open()
            Dim myDataAdapter As New MySqlClient.MySqlDataAdapter(FIELDS & setFrom & setWhere, myCnSQL)
            myDataAdapter.Fill(myDataSet, FROM.ToString)
        Catch ex As Exception
            MsgBox("Error al conectar a la base de datos" + ex.Message)
        Finally
            myCnSQL.Close()
            myDataSet.Dispose()
        End Try
        Return myDataSet
    End Function

    Public Shared Function GetEscalarAsDouble(ByVal statementSQL As String) As Double
        'Funcion que busca y retorna una constante(numero).
        Try
            myCnSQL.Open()
            myCommandSQL = New MySqlClient.MySqlCommand(statementSQL, myCnSQL)
            Return CDbl(myCommandSQL.ExecuteScalar())
        Catch ex As Exception
            ' si hay alguna excepcion retornar CERO en vez de una excepcion
            Return 0
        Finally
            myCnSQL.Close()
            myCommandSQL.Dispose()
        End Try
    End Function
    Public Shared Sub RawQuery(statementSQL As String)
        ABM(statementSQL)
    End Sub
    Public Shared Function VerificarConexion() As Boolean
        'Funcion para verificar si existe conexion a la db, 
        'retorna TRUE si hay conexion y FALSE si no hay conexion.
        Dim estado As Boolean = False
        If estado = False Then
            Try
                myCnSQL.Open()
                myCnSQL.Close()
                estado = True
            Catch ex As Exception
                ' retornar como FALSE si hay algun error en la conexion.
                Return False
            End Try
        End If
        Return estado
    End Function
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        ' Funcion para detectar llamadas redundantes y liberar recursos inutilizados en esta clase.
        If disposing Then
            ' Aqui se deben eliminar los objetos con estado administrado
            '( los objetos que se desean limpiar.)

            myCommandSQL.Dispose()
            myCnSQL.Close()
            myDataSet.Dispose()
        End If

    End Sub
    Private Sub Dispose() Implements IDisposable.Dispose
        ' No cambiar este código. Colocar el código de limpieza en Dispose(disposing As Boolean).
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Private Function FixQueryToProcess(CompleteText As String, WordToFind As String) As Boolean
        Dim Encontrado As Boolean = False
        Dim i As Integer
        i = InStr(1, CompleteText, WordToFind)
        If i > 0 Then
            Encontrado = True
        Else
            Encontrado = False
        End If
        Return Encontrado
    End Function


End Class
