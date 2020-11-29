<!-- #Include Virtual = "/Include/DBCP.asp" -->
<%
Sub closeDB(rs, objDB)
    rs.close
    objDB.close
    
    Set rs = Nothing
    Set objDB = Nothing
End Sub