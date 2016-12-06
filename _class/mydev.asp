<%
class mydev
	private ip_list, isdebug , script_name

	private sub Class_Initialize
		ip_list = "^210.59.162.67|220.134.194.189$"
		isdebug = mynumeric(request("debug"))	
		script_name = Request.ServerVariables("SCRIPT_NAME")
	end sub

	function is_dev_ip
		set reg = new regexp
		with reg
			.pattern = ip_list
			.IgnoreCase = true
		end with

		is_dev_ip = reg.test( Request.ServerVariables ("REMOTE_ADDR"))
	end function

	function debug_js
        debug_js = "nothing.js"
    end function

	function debug
		if isdebug or is_dev_ip then debug = true else debug = false
	end function

	function print ( str_ )
		if debug then
			response.write "<span style='color:#FF0000;'>[debug]</span> <span style='color:#0000FF;'>" & str_ & "</span><br>" & vbcrlf	
		end if
	end function

	
	private sub Class_Terminate
	end sub
end class
%>
