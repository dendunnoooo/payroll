Attribute VB_Name = "Module2"
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
    Set conPayroll = WorkspaceODBC.OpenConnection("", dbDriverNoPrompt, False, "ODBC;Database=DATABASELEOOOO;UID=sa;PWD=pentium;DSN=payroll12345")
End Sub
'--- DAO RECORDSET COSTTRANHEADER FUNCTION
Public Sub openrstPayroll(SelectString As String)
    Set rstPayroll = conEmployee.OpenRecordset(SelectString, dbOpenDynamic, 0, dbOptimistic)
End Sub

