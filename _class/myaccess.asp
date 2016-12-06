<%
class myaccess
	private conn , rs

	private sub Class_Initialize
		set rs = server.createobject("adodb.recordset")

		set conn = server.createobject("adodb.connection")	
		conn.open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("/mission/db/item.mdb")
	end sub

	public function get_city_area ( zip )	
		dim z 
		sql = "SELECT city , area FROM zip WHERE zip3 = '" & zip & "'"
		rs.open sql,conn,1,1
		if not rs.eof then
			redim z(2)
			z(1) = rs("city")
			z(2) = rs("area")
		end if
		rs.close

		get_city_area = z
	end function

end class
%>
