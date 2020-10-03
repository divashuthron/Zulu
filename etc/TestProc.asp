<!--#include Virtual = "/test/common/include/Function.asp" -->
<%
    Set Conn = New clsDBHelper
    Conn.strConnectionString = strDBConnString
    Conn.sbConnectDB

    Dim CommodityName       :   CommodityName       =   fnRF("CommodityName")
    Dim CommodityValue      :   CommodityValue      =   fnRF("CommodityValue")
    Dim CommodityCount      :   CommodityCount      =   fnRF("CommodityCount")

'   SQL = "SELECT * FROM test1"

    SQL =                "INSERT INTO Commodity ( "
    SQL = SQL & vbCrLf & "   CommodityName, CommodityValue, CommodityCount "
    SQL = SQL & vbCrLf & "   ) VALUES ( "
    SQL = SQL & vbCrLf & "   ?, ?, ? "
    SQL = SQL & vbCrLf & "   )"

    Params = Array(_
            Array("@CommodityName",     adVarchar, adParamInput, 20, CommodityName)     _
        ,   Array("@CommodityValue",    adVarchar, adParamInput, 20, CommodityValue)    _
        ,   Array("@CommodityCount",    adVarchar, adParamInput, 20, CommodityCount)    _
    )
    
'   Conn.blnDebug = true
'   Conn.blnViewError = true
    Call Conn.sbExecSQL(SQL, Params)
    Response.Write "success"
    
'   Params = Conn.fnGetArray
'   Arrays = Conn.fnExecSQLGetRows(SQL, Nothing)
%>