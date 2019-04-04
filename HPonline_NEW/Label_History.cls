VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Label_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_strSN As String
Private m_strVer As String
Private m_strPb As String
Private m_strStatue As String
Private m_strPowerCode As String
Private m_strPowerOrigin As String
Private m_strUser As String
Private m_strTime As String

Private conn As New ADODB.Connection
Private rs As New ADODB.Recordset
Private cmd As New ADODB.Command
Private strConnectionString As String
Private strSql As String


Private Sub Class_Initialize()
    strConnectionString = "Provider=SQLOLEDB.1; Data Source=10.11.1.130;Initial Catalog=afg_active_90; User ID=sa; PWD=Flash123"
End Sub

Private Sub Class_Terminate()
    Call Release
End Sub

Private Sub Release()
    If rs.State = ObjectStateEnum.adStateOpen Then
        rs.Close
    End If
    If (conn.State = ObjectStateEnum.adStateOpen) Then
        conn.Close
    End If
    Set cmd = Nothing
    Set rs = Nothing
    Set conn = Nothing
End Sub

Public Function Init(ByVal SerialNumber As String) As Boolean
    If (Len(SerialNumber) <= 0) Then
        Init = False
        Exit Function
    End If
    strSql = "SELECT serial_number,software_version,[5000_status] as status,power_code" & _
            ",power_origin,pb,update_user,modified_time From H3C_PB WHERE serial_number ='" & SerialNumber & "' "

    If conn.State = ObjectStateEnum.adStateClosed Then
        conn.ConnectionString = strConnectionString
        conn.Open
    End If
    
    rs.Open strSql, conn, adOpenKeyset, adLockReadOnly
    If (rs.EOF) Then
        Init = False
        Call Release
        Exit Function
    End If
    
    m_strSN = rs!serial_number
    m_strVer = rs!software_version
    m_strStatue = rs!Status
    m_strPowerCode = rs!power_code
    m_strPowerOrigin = rs!power_origin
    m_strPb = rs!PB
    m_strUser = IIf(IsNull(rs!update_user), "", rs!update_user)
    m_dtTime = rs!modified_time
    
    Init = True
    
End Function

Public Property Get SerialNumber() As String
    SerialNumber = m_strSN
End Property
Public Property Let SerialNumber(ByVal SerialNumber As String)
    m_strSN = SerialNumber
End Property
Public Property Get Version() As String
    Version = m_strVer
End Property
Public Property Let Version(ByVal Version As String)
    m_strVer = Version
End Property
Public Property Get Status() As String
    Status = m_strStatue
End Property
Public Property Let Status(ByVal Status As String)
    m_strStatue = Status
End Property

Public Property Get PowerCode() As String
    PowerCode = m_strPowerCode
End Property
Public Property Let PowerCode(ByVal PowerCode As String)
    m_strPowerCode = PowerCode
End Property
    
Public Property Get PB() As String
    PB = m_strPb
End Property
Public Property Let PB(ByVal PB As String)
    m_strPb = PB
End Property

Public Property Get User() As String
    User = m_strUser
End Property
Public Property Let User(ByVal User As String)
    m_strUser = User
End Property

Public Property Get Time() As String
   Time = m_strTime
End Property
Public Property Let Time(ByVal Time As String)
    m_strTime = Time
End Property

