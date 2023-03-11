Attribute VB_Name = "Module1"
Public CN As New ADODB.Connection
Public RS As New ADODB.Recordset
Public RS1 As New ADODB.Recordset
Public RS2 As New ADODB.Recordset
Public S As String
Public sql As String
Public Sub connect()
    S = " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\BCA\Diet-Hub\Nutri.mdb;Persist Security Info=False"
    CN.Open S
    sql = "select * from Notes"
    RS.Open sql, CN, adOpenDynamic, adLockOptimistic
    
    sql = "select * from Unbalanced"
    RS1.Open sql, CN, adOpenDynamic, adLockOptimistic
    
    sql = "select * from Client"
    RS2.Open sql, CN, adOpenDynamic, adLockOptimistic

End Sub


