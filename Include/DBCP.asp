<%
    Dim DBSIP, DBName, DBID, DBPW, DBCPConfig

    DBSIP = "DESKTOP-NML3K5N"
    DBName = "Zulu"
    DBID = "Zulu"
    DBPW = "zulu"

    DBCPConfig = "Provider=SQLOLEDB; Data Source=" & DBSIP & "; Initial Catalog=" & DBName & "; user ID=" & DBID & "; password=" & DBPW & ";"
%>