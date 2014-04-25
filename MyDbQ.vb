Imports Microsoft.VisualBasic
Public Class MyDbQ
    Private Shared Function q(operation As String, parameter As String, table As String) As String
        'funcion que contruye la query para cada operacion respectiva.
        Dim qv As String = "select " & operation & "(" & parameter & ") from " + table
        Return qv
    End Function
    Public Shared Function Avg(param As String, table As String) As Double
        Return MyDbCore.GetEscalarAsDouble(q("avg", param, table))
    End Function

    Public Shared Function Count(param As String, table As String) As Double
        Return MyDbCore.GetEscalarAsDouble(q("count", param, table))
    End Function

    Public Shared Function Min(param As String, table As String) As Double
        Return MyDbCore.GetEscalarAsDouble(q("min", param, table))
    End Function
    Public Shared Function Max(param As String, table As String) As Double
        Return MyDbCore.GetEscalarAsDouble(q("max", param, table))
    End Function
    Public Shared Function Sum(param As String, table As String) As Double
        Return MyDbCore.GetEscalarAsDouble(q("sum", param, table))
    End Function
End Class

