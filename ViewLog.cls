VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViewLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements LSEXT.IGenericExtension

Private Sub IGenericExtension_Execute(Parameters As LSEXT.IExtensionParametersEx)
          Dim NtlsUser As LSSERVICEPROVIDERLib.NautilusUser
          Dim ServiceProvider As LSSERVICEPROVIDERLib.NautilusServiceProvider
          Dim NtlsCon As LSSERVICEPROVIDERLib.NautilusDBConnection
          Dim con As New ADODB.Connection
        Dim cs As String
          
20        Set ServiceProvider = Parameters("SERVICE_PROVIDER")
30        Set NtlsCon = ServiceProvider.QueryServiceProvider("DBConnection")

          cs = NtlsCon.GetADOConnectionString
          
          If NtlsCon.GetServerIsProxy Then
            cs = "Provider=OraOLEDB.Oracle;Data Source=" & _
            NtlsCon.GetServerDetails & ";User id=/;Persist Security Info=True;"
          End If
          
    
40        con.Open cs
50        con.CursorLocation = adUseClient
60        con.Execute "SET ROLE LIMS_USER"
70        Call ConnectSameSession(con, CDbl(Parameters("SESSION_ID")))

80        Set NtlsUser = ServiceProvider.QueryServiceProvider("User")
          
90        frmLogViewer.strOperatorId = NtlsUser.GetOperatorId
100       Set frmLogViewer.con = con
110       frmLogViewer.Show vbModal
End Sub

Private Sub ConnectSameSession(ByVal aConnection As ADODB.Connection, ByVal aSessionID As Double)
          Dim aProc As New ADODB.Command
          Dim aSession As New ADODB.Parameter
          
120       aProc.ActiveConnection = aConnection
130       aProc.CommandText = "lims.lims_env.connect_same_session"
140       aProc.CommandType = adCmdStoredProc
          
150       aSession.Type = adDouble
160       aSession.Direction = adParamInput
170       aSession.Value = aSessionID
180       aProc.Parameters.Append aSession
          
190       aProc.Execute
200       Set aSession = Nothing
210       Set aProc = Nothing
End Sub

