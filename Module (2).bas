Attribute VB_Name = "Module1"
Public WorkspaceODBC As Workspace
Public conPayroll As Connection
Public rstEmployee As Recordset
Public rstPayroll As Recordset


Public SetupReport
'DAO WORKSPACE FUNCTION
Public Sub openWORKSPACEODBC()
    Set WorkspaceODBC = CreateWorkspace("ODBCWorkpace", "", "Admin", dbUseODBC)
End Sub
'DAO CONNECTION FUNCTION
Public Sub openconPayroll()
    Set conPayroll = WorkspaceODBC.OpenConnection("", dbDriverNoPrompt, False, "ODBC;Database=DATABASELEOOOO;UID=sa;PWD=pentium;DSN=payroll2b")
End Sub
'--- DAO RECORDSET COSTTRANHEADER FUNCTION
Public Sub openrstEmployee(SelectString As String)
    Set rstEmployee = conPayroll.OpenRecordset(SelectString, dbOpenDynamic, 0, dbOptimistic)
End Sub
'--- DAO RECORDSET COSTTRANHEADER FUNCTION
Public Sub openrstPayroll(SelectString As String)
    Set rstPayroll = conPayroll.OpenRecordset(SelectString, dbOpenDynamic, 0, dbOptimistic)
End Sub
