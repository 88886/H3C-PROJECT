VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Unit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_strSN As String
Private m_strWO As String
Private m_strPN As String
Private m_strVer As String
Private m_strPb As String

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
    strSql = "SELECT serial_number, part_number, part_revision, order_number " & _
            "FROM UNIT LEFT JOIN WORK_ORDER ON UNIT.order_key = WORK_ORDER.order_key " & _
            "WHERE serial_number = '" & SerialNumber & "' "

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
    m_strPN = rs!part_number
    m_strWO = rs!order_number
    m_strVer = rs!part_revision
    
    Dim partList As String
    partList = getPartListByOrder(m_strWO)
    If (Len(partList) = 0) Then
        m_strPb = ""
    Else
        m_strPb = GetPBState(partList)
    End If
    Call Release
    
    Init = True
    
End Function



Public Property Get SerialNumber() As String
    SerialNumber = m_strSN
End Property
Public Property Let SerialNumber(ByVal SerialNumber As String)
    m_strSN = SerialNumber
End Property

Public Property Get WorkOrder() As String
    WorkOrder = m_strWO
End Property

Public Property Let WorkOrder(ByVal WorkOrder As String)
    m_strWO = WorkOrder
End Property


Public Property Get PartNumber() As String
    PartNumber = m_strPN
End Property

Public Property Let PartNumber(ByVal PartNumber As String)
    m_strPN = PartNumber
End Property


Public Property Get PartRevision() As String
    PartRevision = m_strVer
End Property

Public Property Let PartRevision(ByVal PartRevision As String)
    m_strVer = PartRevision
End Property

Public Property Get Pb() As String
    Pb = m_strPb
End Property

Public Property Let Pb(ByVal Pb As String)
    m_strPb = Pb
End Property
