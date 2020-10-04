<!--#InClude Virtual="/common/include/Function.asp" -->
<%
    Set objDB = New clsDBHelper
    objDB.strConnectionString = strDBConnString
    objDB.sbConnectDB

    Dim ds_userID           :   ds_userID           = fnRF("ds_user_id")
    Dim ds_userPassword     :   ds_userPassword     = fnRF("ds_user_password")
    Dim ds_saveID           :   ds_saveID           = fnRF("ds_user_is_save")
    Dim ds_autoLogin        :   ds_autoLogin        = fnRF("ds_user_is_auto")

    SQL = "                 SELECT ds_userID, ds_userPassword   "
    SQL = SQL & VBCrLf & "      FROM ds_userTBL                 "
    SQL = SQL & VBCrLf & "      WHERE ds_userID = ?             "
    SQL = SQL & VBCrLf & "      AND ds_userPassword = ?         "

    loginParams = Array(_
            Array("@ds_userID"      , adVarchar, adParamInput, 20, ds_userID      ) _
        ,   Array("@ds_userPassword", adVarchar, adParamInput, 30, ds_userPassword) _
    )

    Call objDB.sbExecSQL(SQL, loginParams)
    Response.Write "Success"

    Set objDB = Nothing
    response.end
%>
