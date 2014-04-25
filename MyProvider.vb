Option Explicit On
Option Strict On
Imports Microsoft.VisualBasic

Public Class MyProvider
    'Metodo creado para definir la ruta de acceso a la base de datos, 
    'Se ha encapsulado en una propiedad de acceso en modo de lectura para garantizar 
    'la seguridadde la ruta y evitar que se pueda modificar al invocar el metodo.
    '/**
    '*
    '* @author: daMgeL
    '* @date:   03-03-2014:2:04PM
    '*
    '/**
    Public Shared ReadOnly Property Ruta() As String

        Get
            Dim dbruta As String = "server=localhost; user id=root; password=destiny; persist security info=True; database=sakila"
            Return dbruta
        End Get


    End Property
End Class

