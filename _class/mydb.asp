<%
class mydb

	public is_debug
	private sql_query, host, db_name , user , passwd

	private is_write , sd , table_name
	private conn , rs0, q
	private this_pkey
	
	public rs, last_sn, item, has_data
	

	private sub Class_Initialize
		set rs = server.createobject("adodb.recordset")
		set rs0 = server.createobject("adodb.recordset")
		last_sn = 0
		this_pkey = "sn"
		set item = createobject("scripting.dictionary")
	end sub

	public default function construct ( h , n , u , p )
		host = h	
		db_name = n
		user = u
		passwd = p

		set conn = server.createobject("adodb.connection")
		conn.open "Provider=SQLOLEDB;Data Source=" & host &";database=" & db_name & ";User ID=" & user & ";Password=" & passwd & ";"
		set construct = me
	end function
	
	private sub Class_Terminate
		set item = nothing
	end sub
	
	public function set_pkey ( str_ )
		this_pkey = str_
	end function
	
	public function get_datas ( tb_name , pkey )
		pkey = mynumeric(pkey)
		item.removeall
		items = get_table_columns ( tb_name )

		sql = "SELECT * FROM " & tb_name & " WHERE " & this_pkey & " = " & pkey
		query sql,false
		if not rs.eof then
			for each f in items
				item.add f , rs(f)
			next
			has_data = true
		else
			has_data = false
		end if

	end function
	
	public function tojson 
		for each key in item.keys 
			json_d(key) = item(key)
		next
	end function

	private function get_table_columns ( tb_name )
	    t_a = split(tb_name,".")
	    if ubound(t_a) = 2 then
    	    conn.DefaultDatabase = t_a(0)
       		sql = "SELECT name FROM syscolumns WHERE id = object_id('" & t_a(2) & "')"
	    else
        	sql = "SELECT name FROM syscolumns WHERE id = object_id('" & tb_name & "')"
	    end if

   		dim c

		query sql,false

   	 	if not rs.eof then redim c(rs.recordcount - 1)

	    c_i = 0
   		while not rs.eof
        	c(c_i) = rs("name")
 	 		c_i = c_i + 1
       		rs.movenext
    	wend
	    rs.close

    	get_table_columns = c
	end function

	public function get_last_sn ( tb_name, pkey )
		sql = "select max(" & pkey & ") as last_sn from " & tb_name 
		query sql,false
		get_last_sn = get_data("last_sn")
	end function

	public function query ( s , w) 
		sql_query = s
		is_write = w
		print 

		if is_write then 
			mode = 3 
			get_table_name
			get_table_attr
		else 
			mode = 1
		end if
		rs.open sql_query,conn,1,mode

		if table_name = sql_query and mode = 3 then
			rs.addnew 
		end if

'		query = rs
	end function

	public function get_data ( key )
		if not rs.eof then get_data = rs(key) 
	end function

	public sub set_data ( key , value )
'		print_str ( key)
		if sd.exists(key) then
			t_a = split(sd.item(key),":")
			select case t_a(0)
				case "int","bit"
					value = mynumeric(value)
				case "varchar","nvarchar","char"
					value = left(value,t_a(1))
			end select
			rs(key) = value
		end if
	end sub
	
	public sub reset_db ( tb_name )
		sql_query = "delete from " & tb_name & " ; dbcc checkident( " & tb_name & " , reseed , 0);"
		print
		conn.execute(sql_query)
	end sub

	public sub exec ( str_ )
		sql_query = str_
		print
		conn.execute(sql_query)
	end sub

	private sub get_table_name
		str_ = lcase(sql_query)
		f_p = instr(str_,"from")
		w_p = instr(str_,"where")
		if f_p > 0 and w_p > f_p then
			t_str = mid(str_,f_p,w_p - f_p) 
		elseif f_p > 0 and w_p = 0 then
			t_str = mid(str_,f_p) 
		else
			t_str = str_
		end if

		table_name = trim(replace(t_str,"from",""))

	end sub

	sub print
		if is_debug then
			response.write "<p style='color:#FF0000;'>" & sql_query & "</p>"
		end if
	end sub


	public sub close
		set sd = nothing
		rs.close
	end sub

	public sub commit
		rs.update
	end sub

	public sub addnew
		rs.addnew
	end sub

	public sub delete
		rs.delete
	end sub

	public sub print_str ( str_ )
		response.write "<p style='color:#0000FF;'>" & str_ & "</p>"	
	end sub

	private sub get_table_attr 
		if len(table_name) = 0 then exit sub
		sql = "select dbo.syscolumns.name AS sColumnsName, dbo.syscolumns.prec AS iColumnsLength, dbo.systypes.name + '' AS sColumnsType FROM dbo.sysobjects INNER JOIN dbo.syscolumns ON dbo.sysobjects.id = dbo.syscolumns.id INNER JOIN dbo.systypes ON dbo.syscolumns.xusertype = dbo.systypes.xusertype WHERE (dbo.sysobjects.xtype = 'U') and dbo.sysobjects.name = '" & table_name & "'"

		rs0.open sql,conn,1,1
		if not rs0.eof then
			set sd = server.createobject("scripting.dictionary")
		end if

		while not rs0.eof
			val_ = rs0("sColumnsType") & ":" & rs0("iColumnsLength") 
			key_ = rs0("sColumnsName")
			sd.add key_ , val_
			rs0.movenext
		wend
		rs0.close
	end sub

	public function chkSameFieldValue ( tb_name , fname , value )
		chkSameFieldValue = false	
		sql = "SELECT " & fname & " FROM " & tb_name & " WHERE " & fname & " = '" & value & "'"
		query sql,false
		if rs.recordcount > 0 then chkSameFieldValue = true
		rs.close	
	end function

	public function gen_query ( tb_name , fname , value )
		gen_query = "SELECT * FROM " & tb_name & " WHERE " & fname & " = '" & value & "'"
	end function

	public function del_data ( tb_name , pkey )
		if not isnumeric(pkey) then exit function

		sql = "SELECT * FROM " & tb_name & " WHERE sn = " & pkey
		
		query sql,true

		if not rs.eof then
			rs.delete
			del_data = true
		else
			del_data = false
		end if
		rs.close
		
	end function

end class
%>
