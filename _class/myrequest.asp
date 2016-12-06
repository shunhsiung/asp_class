<%
class myrequest
	public count , items , err_msg 

	public function checked_c ( item , value )
		checked_c = ""
		if items.exists(item) and instr(" " & items(item), value) > 0  then checked_c = " checked "
	end function

	public function checked ( item )
		checked = ""
		if items.exists(item) and items(item) then checked = " checked "
	end function

	public function get_items()
		str_t = ""
		for each item in items.keys
			str_t = str_t & "{{" & item & "::" & items.item(item) & "}}"
		next
		get_items = str_t
	end function

	public sub set_items_data ( str_t , clean)
		if clean then items.removeall 

		if len(str_t) = 0 then exit sub
		jlist = split(str_t,"}}")
		for each item in jlist
			item = replace(item,"{{","")			
			it = split(item,"::")
			if isarray(it) and ubound(it) = 1 then
				if not items.exists(it(0)) then
					items.add it(0), it(1)	
				else
					items(it(0)) = it(1)
				end if
			end if
		next
	end sub

	public sub chk_item(item , value ) 
		if not items.exists(item) then
			items.add item ,value
		end if
	end sub
	
	public sub set_default ( item , value )
		if items.exists(item) then
			if len(items(item)) > 0 then
				exit sub
			end if	
		end if
		set_items item ,value
	end sub

	private sub set_items ( item , value )
		if items.exists(item) then
			items(item) = value
		else
			items.add item ,value
		end if
	end sub

	public sub remove_item ( n_a )
		for each item in n_a
			if items.exists(item) then items.remove item
		next
	end sub

	public sub chkmyperson_id ( item , msg )
		if items.exists(item) then
			if not chkpersonid (items(item)) then err_msg = err_msg & msg  
		end if
	end sub

   public sub chkmyperson_id2 ( item , msg )
        if items.exists(item) then
            if not chkpersonid (items(item)) then
                if len(items(item)) < 8 then
                    err_msg = err_msg & msg
                end if
            end if
        end if
    end sub

	public sub chkmymobile ( item , msg )
		if items.exists(item) then
			items(item) = replace(items(item),"-","")
			if not chkmobile (items(item)) then err_msg = err_msg & msg  
		end if
	end sub

	public sub chkmydate ( item , msg )
		if items.exists(item) then
			if not isdate(items(item)) then err_msg = err_msg & msg
		end if
	end sub

	public sub chkmyemail ( item , msg )
		if items.exists(item) then
			if not chkemail (items(item)) then err_msg = err_msg & msg  
		end if
	end sub

	public sub chkpkey ( item , msg)
		if not items.exists(item) then 
			err_msg = err_msg & msg
		else
			if items(item) = 0 then err_msg = err_msg & msg	
		end if
	end sub

	public sub chkempty_m ( n_a , msg )
		n_a_max = ubound(n_a) + 1
		empty_count = 0
		for each item in n_a
			if items.exists(item) then		
				if len(items(item)) = 0 then empty_count = empty_count + 1
			else
				empty_count = empty_count + 1
			end if
		next

		if empty_count = n_a_max then 
			err_msg = err_msg & msg
		end if
	end sub

	public sub chkempty ( n_a )
'		err_msg = ""
		for each item in n_a
			item_a = split(item,":")
			if items.exists(item_a(0)) then 
				if len(items(item_a(0))) = 0 then err_msg = err_msg & item_a(1)
			else
				err_msg = err_msg & item_a(1)
			end if
		next
	end sub

	public sub set_space ( n_a )	
		for each item in n_a
			if not items.exists(item) then items(item) = ""
		next
	end sub

	public sub set_numeric ( n_a )
		for each item in n_a
			if items.exists(item) then items(item) = mynumeric(items(item))
		next
	end sub

	Private Sub Class_Initialize 
		set items = server.createobject("Scripting.Dictionary")
		count = request.form().count
		if count > 0 then
			for each item in request.form
				if len(trim(item)) > 0 and item <> "undefined" and item <> "args" then
					set_items item , request.form(item)
				end if
			next
		end if

        args = request.form("args")

        if len(args) > 3 then
            t_a = split(args,",")
            for each t in t_a
                t1 = split(t,":")
                if ubound(t1) = 1 then
					set_items t1(0) , t1(1)
                end if
            next
        end if
        
	End Sub

	Private Sub Class_Terminate
	End Sub

end class
%>
