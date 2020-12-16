<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/EXPRESS.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_EXPRESS_STRING
    MM_editCmd.CommandText = "UPDATE dbo.Send SET Seno = ?, Sno = ?, Bno = ?, Cno = ?, Gno = ?, sitime = ?, sotime = ?, setype = ? WHERE Seno = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("Seno"), Request.Form("Seno"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("Sno"), Request.Form("Sno"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("Bno"), Request.Form("Bno"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("Cno"), Request.Form("Cno"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("Gno"), Request.Form("Gno"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 135, 1, -1, MM_IIF(Request.Form("sitime"), Request.Form("sitime"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 135, 1, -1, MM_IIF(Request.Form("sotime"), Request.Form("sotime"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 201, 1, 20, Request.Form("setype")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "商家查看配送表.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("Seno") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("Seno")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_EXPRESS_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.Send WHERE Seno = ? ORDER BY Seno ASC" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 5, 1, -1, Recordset1__MMColParam) ' adDouble

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
</head>

<body>
<form action="<%=MM_editAction%>" method="POST" name="form1" id="form1">
  <table align="center">
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Seno:</td>
      <td><input type="text" name="Seno" value="<%=(Recordset1.Fields.Item("Seno").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Sno:</td>
      <td><input type="text" name="Sno" value="<%=(Recordset1.Fields.Item("Sno").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Bno:</td>
      <td><input type="text" name="Bno" value="<%=(Recordset1.Fields.Item("Bno").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Cno:</td>
      <td><input type="text" name="Cno" value="<%=(Recordset1.Fields.Item("Cno").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Gno:</td>
      <td><input type="text" name="Gno" value="<%=(Recordset1.Fields.Item("Gno").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Sitime:</td>
      <td><input type="text" name="sitime" value="<%=(Recordset1.Fields.Item("sitime").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Sotime:</td>
      <td><input type="text" name="sotime" value="<%=(Recordset1.Fields.Item("sotime").Value)%>" size="32" /></td>
    </tr>
    <tr>
      <td nowrap="nowrap" align="right" valign="top">Setype:</td>
      <td valign="baseline"><table align="left">
        <tr>
          <td><input type="radio" value="" name="setype" />
            邮政快递 </td>
          <td></td>
        </tr>
        <tr>
          <td><input type="radio" value="" name="setype" />
            中通快递 </td>
          <td></td>
        </tr>
        <tr>
          <td><input type="radio" value="" name="setype" />
            申通快递 </td>
          <td></td>
        </tr>
        <tr>
          <td><input type="radio" value="" name="setype" />
            东风快递 </td>
          <td></td>
        </tr>
        <tr>
          <td><input type="radio" value="" name="setype" />
            其他 </td>
          <td></td>
        </tr>
      </table></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">&nbsp;</td>
      <td><input type="submit" value="更新记录" /></td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1" />
  <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("Seno").Value %>" />
</form>
<p>&nbsp;</p>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
