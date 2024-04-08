VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CreateLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public con As ADODB.Connection
Public session As Long

Public Sub InsertLog(SdgId As Long, ApplicationCode As String, Description As String)
10        Call con.Execute("insert into lims_sys.sdg_log " & _
                  "(SDG_ID, TIME, APPLICATION_CODE, SESSION_ID, DESCRIPTION) " & _
                  "values (" & SdgId & ",sysdate,'" & ApplicationCode & "'," & _
                  session & ",'" & Description & "')")
End Sub
