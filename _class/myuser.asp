<%
class myuser
    private web_ticket
    private item
    private table
	private account_file

    public err_msg
    public islogin
    public success
    public user_ip

    private sub Class_Initialize 
        web_ticket = config.web_ticket
        set item = createobject("scripting.dictionary")
        table = "zion.dbo.cowork" 
        user_ip = Request.servervariables("REMOTE_ADDR")
        call init()
    end sub

    private sub Class_Terminate
        set item = nothing     
    end sub
     
    public function logout
        empty_cookies(web_ticket)
    end function

    public function login
        account = chkpost(trim(request("account")))
        password = trim(request("password"))
        
		isdevloper = false
		success = false

		if false and dev.is_dev_ip then
			account = "sshsiung"
			password = "swh5912b2"
			isdeveloper = true
		end if

		account_a = get_account_list 

		for each item in account_a
        	if instr(item,":" & account & ":") > 0 then
            	i_a = split(item,":")
            	if isarray(i_a) and ubound(i_a) = 4 then
                	sn = i_a(0)
	                i_account = i_a(1)
   		            i_password = i_a(2)
        	        cname = i_a(3)
           		    level = i_a(4)
                	if password = i_password or isdeveloper then
'                    	t_l = (clng(level) + 1925) * clng(sn)
                    	t = sn & ":" & i_account & ":" & cname & ":" & web_ticket 
                    	c_t = len(t) * 97
                    	a_ticket = t & ":" & cstr(c_t)
                    	call set_cookies_value(web_ticket,a_ticket,3)
						success = true
	   	             end if
   				end if
	        end if
   		next

		if not success then err_msg = "帳號或密碼錯誤!!"

    end function

	private function get_account_list
		str_ = file2str("/mission/_priv/account.db")
		account_a = split(str_,chr(10))

		get_account_list = account_a
	end function

    public function init
        a_ticket = get_cookies_value(web_ticket)  
        islogin = false
        if len(a_ticket) = 0 then exit function

        if len(a_ticket) > (9 + len(web_ticket)) then
            ts = split(a_ticket,":")
            ts_max = ubound(ts)
            if ts_max = 4 then
                if trim(ts(3)) = trim(web_ticket) then
                    if cint(ts(4)) = len(replace(a_ticket,":" & ts(4),"")) * 97 then  
                        islogin = true
                        item.add "sn" , ts(0)
                        item.add "account" , ts(1)
                        item.add "cname" , ts(2)
						if ts(0) = 1 then
							item.add "level" , "admin"
						else
							item.add "level" , "user"
						end if
'						call get_data (ts(0))
                    end if
                end if
            end if
        end if
    end function

	private sub get_data ( sn )
		sn = mynumeric(sn)
        sql = "select * from " & table & " where sn = '" & sn & "'"
		rs.open sql,conn,1,1
		if not rs.eof then				
			last_ip = rs("last_ip")
			last_login = rs("last_login")
			groups = rs("groups")
			p_basic = rs("P_Basic")
			p_mc21 = rs("P_Mc21")
			p_goods = rs("P_Goods")
		end if
		rs.close

		item.add "groups", groups	
		item.add "p_basic" , p_basic
		item.add "p_mc21" , p_mc21
		item.add "p_goods" , p_goods
		item.add "last_ip" , last_ip
		item.add "last_login" , last_login
	end sub

	public function data(key)
		if item.exists(key) then
			data = item(key)
		else
			data = ""
		end if
	end function

end class
%>
