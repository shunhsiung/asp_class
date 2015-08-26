<%
class mydev
	private ip_list, isdebug , script_name
	private count_i

	private sub Class_Initialize
		ip_list = "^192.168.1.1|192.168.3.1$"
		isdebug = mynumeric(request("debug"))	
		script_name = Request.ServerVariables("SCRIPT_NAME")
		count_i = 0
	end sub

	public default function construct ( list )
		ip_list = list

		set construct = me
	end function

	function is_dev_ip
		is_dev_ip = chkregexp ( Request.ServerVariables ("REMOTE_ADDR") , ip_list )			
	end function

	function debug
		if isdebug or is_dev_ip then debug = true else debug = false
	end function

	function debug_js
		debug_js = "nothing.js"
	end function

	function print ( str_ )
		if debug then
			response.write "<p style='color:#FF0000;'>" & str_ & "</p>"	
		end if
	end function

	function print_count ( str_ )
		count_i = count_i + 1
		print count_i & " " & str_
	end function
	
	function print_count_end
		count_i = 0
	end function

	private sub Class_Terminate
	end sub
end class
%>
