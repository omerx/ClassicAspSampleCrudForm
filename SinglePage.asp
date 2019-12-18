<%
Call SubAdmin

ID=tc(Request("ID"),"i")
VarcharCol=tc(Request("VarcharCol"),"v")
IntegerCol=tc(Request("IntegerCol"),"i")
DateCol=tc(Request("DateCol"),"d")

set vb=Server.CreateObject("Adodb.Connection")
vb.Open "Provider=SQLNCLI11;Server=.;Database=Database;Uid=sqlUser;Pwd=sqlPass;Workstation ID=ClassicAspForever;"


p=Request("p") 'Process
Select Case p
Case "Form"				: Call SubForm
Case "Add"				: Call SubAdd
Case "Edit"				: Call SubEdit
Case "Delete"			: Call SubDelete
Case Else               : Call SubList
End Select


Sub SubAdmin
    if Request("pass")="MyPassword" then    : Session("admin")="true"
    if not Session("admin")="true" then     : Response.Write "Not Admin" : Response.End
End Sub


Sub SubList
Call SubNav
    set rs=vb.execute("Select * From [Table]",rc,1)
        While not rs.eof
            Response.Write "<p><b>"&rs("ID")&"=</b>"&rs("VarcharCol")&"-"&rs("IntegerCol")&"-"&rs("DateCol")&" | <a href='?p=Form&ID="&rs("ID")&"'>Edit</a>  | <a href='?p=Delete&ID="&rs("ID")&"'>Delete</a> </p>"
        rs.MoveNext:Wend
    set rs=Nothing
End Sub'SubList



Sub SubAdd
Call SubNav
    set rs=vb.execute("INSERT INTO [Table] ([VarcharCol],[IntegerCol],[DateCol]) values ('"&VarcharCol&"','"&IntegerCol&"','"&DateCol&"')",rc,1)
        if rc>0 then : Response.Write "Record Added."
    set rs=Nothing
End Sub'SubAdd


Sub SubEdit
Call SubNav
    set rs=vb.execute("UPDATE [Table] SET [VarcharCol]='"&VarcharCol&"',[IntegerCol]='"&IntegerCol&"',[DateCol]='"&DateCol&"' WHERE ID='"&ID&"'",rc,1)
        if rc>0 then : Response.Write rc &" record Updated."
    set rs=Nothing
End Sub'SubEdit

Sub SubDelete
Call SubNav
    set rs=vb.execute("DELETE [Table] WHERE ID='"&ID&"'",rc,1)
        if rc>0 then : Response.Write rc &" record Deleted."
    set rs=Nothing
End Sub'SubDelete

Sub SubForm
Call SubNav
    'Data Bind
    set rs=vb.execute("Select  ID,VarcharCol,IntegerCol,convert(varchar(10),DateCol,121) as DateCol From [Table] where ID='"&ID&"'",rc,1)
        if rs.eof then
            p="Add"
        Else
            p="Edit"
            ID=rs("ID")
            VarcharCol=rs("VarcharCol")
            IntegerCol=rs("IntegerCol")
            DateCol=rs("DateCol")
        End If
    set rs=Nothing

    Response.Write "<form method='post' action='?p="&p&"'>"&_
                    "<input type='number' name='ID' value='"&ID&"'>"&_
                    "<input type='text' name='VarcharCol' value='"&VarcharCol&"'>"&_
                    "<input type='number' name='IntegerCol' value='"&IntegerCol&"'>"&_
                    "<input type='date' name='DateCol' value='"&DateCol&"'>"&_
                    "<input type='submit'>"&_
                    "</form>"
End Sub'SubForm

Sub SubNav
Response.Write "<a href='?p=List'>List</a> | <a href='?p=Form'>Add & Edit Form</a> <hr>"
End Sub

vb.close
set vb=Nothing

Function tc(d,t) 'TypeControlFunction
    if t&""="i" and isNumeric(d)=true then : tc = d
    if t&""="d" and isdate(d)=true    then : tc = d
    if t&""="v" then :                       tc = Replace(d&"","'","''",1,-1,1)
End Function'tc
%>