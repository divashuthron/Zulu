<%@Language="VBScript" CODEPAGE="65001" %>
<?xml version="1.0" encoding="utf-8" ?>
<!-- #Include Virtual = "/Include/CustomFunction.asp" -->
<%
    Dim rs, objDB, SQL, strResult, ReturnMSG

    Dim ID : ID = Request("ID")
    Dim PW : PW = Request("Password")

    Set objDB = Server.CreateObject("ADODB.Connection")
    objDB.open(DBCPConfig)

    SQL = SQL & VbCrLf & "Select"
    SQL = SQL & VbCrLf & "	Case When ID = '" & ID & "' And Password = '" & PW & "' Then 'Y'"
    SQL = SQL & VbCrLf & "		Else 'N'"
    SQL = SQL & VbCrLf & "	End as ConfirmResult"
    SQL = SQL & VbCrLf & "From Member"

    Set rs = Server.CreateObject("ADODB.RecordSet")
    rs.Open SQL, objDB

    If Not rs.eof or Not rs.bof Then
        If rs("ConfirmResult") = "Y" Then
            strResult = "Success"
            ReturnMSG = ""
        Else
            strResult = "Not Found"
            ReturnMSG = ""
        End If
    End If

    closeDB(rs, objDB)
%>
<Zulu>
    <datas>
        <data>
            <strResult><%= strResult %></strResult>
            <ReturnMSG><%= ReturnMSG %></ReturnMSG>
        </data>
    </datas>
</Zulu>