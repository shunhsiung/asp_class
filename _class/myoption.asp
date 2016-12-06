<%
class myoption

	private rs
	private table 

	private sub Class_Initialize
		table = "seven.dbo.tb_option"
		set rs = createobject("adodb.recordset")	
	end sub

	private sub Class_Terminate
		set rs = nothing
	end sub

	private function query_value( sql )
		rs.open sql,conn,1,1
		if not rs.eof then
			query_value = rs("value_")
		else
			query_value = ""
		end if
		rs.close
	end function	

	public function get1 ( key )
		sql = "select value_ from " & table	 & " where key_ = '" & key & "'"
		get1 = query_value (sql)
	end function

	public function get2 ( key , key2)
		sql = "select value_ from " & table	 & " where key_ = '" & key & "' and key2 = '" & key2 & "'"
		get2 = query_value (sql)
	end function

	public sub save1 ( key , value )	
		call save2 ( key , "", value)
	end sub

	public sub save2 ( key , key2 , value )
		sql = "select * from " & table & " where key_ = '" & key & "' and key2 = '" & key2 & "'"	

		rs.open sql,conn,1,3
		if rs.recordcount = 0 then
			rs.addnew
			rs("key_") = key
			rs("key2") = key2
		end if

		rs("value_") = value
		rs.update
		rs.close
	end sub	

end class
%>
