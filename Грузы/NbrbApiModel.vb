
Imports System.Collections.Generic
Imports System.ComponentModel.DataAnnotations
Imports System.Linq
Imports System.Web

Public Class NbrbApiModel
    Public Class Currency
        <Key>
        Public Property Cur_ID As Integer
        Public Property Cur_ParentID As Nullable(Of Integer)
        Public Property Cur_Code As String
        Public Property Cur_Abbreviation As String
        Public Property Cur_Name As String
        Public Property Cur_Name_Bel As String
        Public Property Cur_Name_Eng As String
        Public Property Cur_QuotName As String
        Public Property Cur_QuotName_Bel As String
        Public Property Cur_QuotName_Eng As String
        Public Property Cur_NameMulti As String
        Public Property Cur_Name_BelMulti As String
        Public Property Cur_Name_EngMulti As String
        Public Property Cur_Scale As Integer
        Public Property Cur_Periodicity As Integer
        Public Property Cur_DateStart As System.DateTime
        Public Property Cur_DateEnd As System.DateTime
    End Class
    Public Class Rate
        <Key>
        Public Property Cur_ID As Integer
        Public Property _Date As DateTime
        Public Property Cur_Abbreviation As String
        Public Property Cur_Scale As Integer
        Public Property Cur_Name As String
        Public Property Cur_OfficialRate As Decimal?
    End Class

    Public Class RateShort
        Public Property Cur_ID As Integer
        <Key>
        Public Property _Date As System.DateTime
        Public Property Cur_OfficialRate As Decimal?
    End Class



End Class
