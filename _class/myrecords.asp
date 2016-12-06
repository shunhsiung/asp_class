<%
class myrecords

	public view , del , edit , page , desc , sortitem , actfile, search_string, other_args

	private page_size , sql , pkey , record_count , page_max , desc_, si , si_max 

	public sub gen_records ( fi, where_ )
		sql_i = ""
		si = fi
		si_max = ubound(si)
		for i = 1 to si_max
			if len(si(i,0)) > 0 then
				sql_i = sql_i & si(i,0) & ","	
			end if
		next


		' 特別搜尋

		sp_where = ""

		if len(si(0,2)) > 0 and len(si(0,3)) > 0 and len(fi(0,4)) > 0 then
			sp_value = ""
			select case fi(0,4)
				case "cookies"
					sp_value = request.cookies(si(0,3))
				case "session"
					sp_value = session(si(0,3))
			end select

			if len(sp_value) = 0 then
				sp_where = " " & si(0,2) & " = 'foo'"
			else
				sp_where = " " & si(0,2) & " = '" & sp_value & "'"
			end if
		end if

		sql = "select " & sql_i & pkey & " from " & si(0,0)
		
		if len(where_) > 0 then sql = sql & " where " & where_

		if len(sp_where) > 0 then
			if len(where_) > 0 then
				sql = sql & " and (" & sp_where & ")"
			else
				sql = sql & " where (" & sp_where & ")"
			end if
		end if
		
		if len(sortitem) > 0 then sql = sql & " order by " & sortitem & " " & desc
		
		rs.open sql,conn,1,1	
		rs.pagesize = page_size
		record_count = rs.recordcount
		page_max = int(record_count / page_size )	

		if (record_count mod page_size ) > 0 then
			page_max = page_max + 1
		end if

		if clng(page) > page_max then page = page_max

		if not rs.eof then rs.absolutepage = page

	end sub
	
	private sub show_action_box()
%>
<div class='row'>
	<div class='well'>	
	<button class='btn btn-primary' data-toggle='modal' data-target='#modal-from-search'><%=jm("__SEARCH__")%></button>
<% if len(session("SText")) > 0 then %>
	<input type=hidden name=actfile value='<%=actfile%>'>
	<input type=hidden name=sortitem value='<%=sortitem%>'>
	<input type=hidden name=desc value='<%=desc%>'>
<%
	for each o_a1 in split(other_args,"?")
		o_a = split(o_a1,"=")
		if isarray(o_a) and ubound(o_a) = 1 then
%>
	<input type=hidden name='<%=o_a(0)%>' value='<%=o_a(1)%>'>
<%
		end if
	next
%>
	<button class='btn' id=clear_search ><%=jm("__CLEAN_SEARCH__")%></button>
<% end if %>
<% if del then %>
	<button class='btn btn-danger' id=del_button><%=jm("__DELETE__")%></button>
	<button class='btn btn-primary' id=all_button><%=jm("__ALL__")%></button>
<% end if %>
	<button class='btn' id=download_button ><%=jm("__DOWNLOAD__")%></button>
	</div>
</div>

<div id='modal-from-search' class='modal fade hide'>
	<form action='<%=actfile%>' method=post id=search_submit>
	<input type=hidden name=sortitem value='<%=sortitem%>'>
	<input type=hidden name=desc value='<%=desc%>'>
<%
	for each o_a1 in split(other_args,"?")
		o_a = split(o_a1,"=")
		if isarray(o_a) and ubound(o_a) = 1 then
%>
	<input type=hidden name='<%=o_a(0)%>' value='<%=o_a(1)%>'>
<%
		end if
	next
%>
	<div class='modal-header'>
		<h3><%=search_string%></h3>
	</div>
	<div class='modal-body'>
		<input type=text name='stext' value='<%=session("Stext")%>' class='span5'>
	</div>
	<div class='modal-footer'>
		<button class='btn' id='cancel_button'><%=jm("__CANCEL__")%></button>
		<button class='btn btn-primary' id='search_button'><%=jm("__SEARCH__")%></button>
	</div>
	</form>
</div>
<%
	end sub

	public sub show_records_file( f_a )
		call show_menu()

		sn_t = f_a(0)
		cname_t = f_a(1)
		thumb_t = f_a(2)
		url_t = f_a(3)
		size_t = f_a(4)

		count_i = (page -1 ) * page_size
%>
<div class='row'>
	<ul class='thumbnails'>
<%
		for p_i = 1 to page_size
			if rs.eof then exit for
%>
		<li class='span3' id='t<%=rs(sn_t)%>'>
			<div class='thumbnail'>
				<img src='<%=rs(thumb_t)%>' class='file_click' data-url='<%=rs(url_t)%>'>
				<h5><%=rs(cname_t)%></h5>
				<span style='font-size:12px;'><%=formatnumber(rs(size_t),0)%> bytes</span>
			</div>
		</li>
<%
			rs.movenext
		next	
%>
	</ul>
</div>
<%
		call show_menu()
	end sub

	public sub show_records ()
		call show_menu()
		call show_action_box()
		if desc = "desc" then desc_ = "asc"
%>
<div class='row'>
<table class='table table-bordered table-striped'>
<%
		call show_title() 
		count_i = (page -1 ) * page_size

%>
	<tbody>
<%
		for p_i = 1 to page_size
			if rs.eof then exit for
			data_sn = rs(pkey)
%>
	<tr id=t<%=data_sn%>>			
		<td><%=count_i+p_i%>.
<%
			if view then 
%>
		<span class='divider'> / </span><button class='btn view_button' value='<%=data_sn%>'><%=jm("__VIEW__")%></button>
<%
			end if

			if edit then 
%>
		<span class='divider'> / </span><button class='btn edit_button' value='<%=data_sn%>'><%=jm("__EDIT__")%></button>
<%
			end if

			if del then 
%>
		<span class='divider'> / </span><input  type=checkbox name=dkey value='<%=data_sn%>'>
<%
			end if
%>
		</td>		
<%
			for j = 1 to si_max
				if si(j,2) = "show" then call show_item(j)
			next
%>
	</tr>
<%
			rs.movenext
		next
%>
	</tbody>
</table>
</div>
<%
		rs.close
		call show_menu()
		call show_foot()
	end sub

	private sub show_foot()

		str_msg = jm("__FOOT_MSG__")
		str_msg = replace(str_msg,"__PAGE_MAX__",page_max)
		str_msg = replace(str_msg,"__RECORD_COUNT__",record_count)
		str_msg = replace(str_msg,"__PAGE__",page)
		str_msg = replace(str_msg,"__PAGE_SIZE__",page_size)
%>
<div class='row'>	
<div class='alert alert-info'><%=str_msg%></div>
</div>
<%		
	end sub

	private sub show_item ( j )
		data_t = rs(si(j,0))
		response.write "<td>"
		select case lcase(si(j,4))
			case "number"	
				data_t = mynumeric(data_t)
				response.write formatnumber(data_t,0)
			case "yes"
				select case si(j,5)
					case "0"
						show_yes_image data_t
				end select
			case "cdate"
				if isDate(data_t) then response.write datevalue(data_t)
			case "thumb2"
				t_a = split(si(j,5),",")
				if isarray(t_a) and ubound(t_a) = 1 then	
					img_t = rs(t_a(1)) & rs(t_a(0))	
					response.write "<img src='" & img_t & "' class='img-polaroid'><p>" & data_t & "</p>"
				else
					response.write data_t
				end if
			case "thumb"
				if len(rs(si(j,5))) > 5 then
					response.write "<img src='" & rs(si(j,5)) & "' class='img-polaroid'><p>" & data_t & "</p>"
				else					
					response.write data_t
				end if
			case "file_image"
				f_a = split(data_t,"/")
				if isarray(f_a) then
					response.write "<img src='img/file/file_extension_" & f_a(1) & ".png' border=0>"
				end if
			case "image"
				base_url = si(j,5)	
				response.write "<img src='" & base_url & data_t & "' class='img-polaroid span2'>"
			case "custom"	
'				select case si(j,5)
'					case "button"
'						b_a = split(si(j,6),",")
'						response.write "<button class='btn " & b_a
'						response.write func_20120427a ( data_t )
'				end select

			case else
				response.write data_t
		end select
		response.write "</td>"

	end sub

	private sub show_title()
%>
	<thead>
	<tr >
	<td ></td>
<%

	for f_i = 1 to si_max
		if si(f_i,2) = "show" then
			if desc = "desc" then
				arrow_img = "<i class='icon-arrow-up'></i>"
			else
				arrow_img = "<i class='icon-arrow-down'></i>"
			end if

			response.write "<th >" & si(f_i,1 ) & "<a href='" & actfile & "?sortitem=" & si(f_i,0) & "&desc=" & desc_ & "&" & other_args & "'>" & arrow_img & "</a></th>"
		end if
	next

'	item_a_max = ubound(item_a)
'	for i_i = 1 to item_a_max
'		if len(item_a(i_i,1)) > 0 then
'			response.write "<th class='td_top'><input type=button id=" & item_a(i_i,2) & "_title value='" & item_a(i_i,3) & "'></th>"
'		end if
'	next

%>
	</tr>
	</thead>
<%
	end sub

	private sub show_menu ()
		NowPage = page
		PageMax = page_max

		if PageMax <= 1 then exit sub

		other_var = "&sortitem=" & sortitem & "&desc=" & desc
		if len(other_args) > 0 then other_var = other_var & "&" & other_args

		if NowPage <> 1 then
			first_url = actfile & "?page=1" & other_var
			prev_url = actfile & "?page=" & NowPage - 1 & other_var
		else
			first_url = "#"
			prev_url = "#"
			prev_disabled = " disabled "
		end if

		if NowPage <> PageMax then
			next_url = actfile & "?page=" & NowPage + 1 & other_var
			last_url = actfile & "?page=" & Pagemax & other_var
		else
			next_url = "#"
			last_url = "#"
			next_disabled = " disabled "
		end if

    	if PageMax < 10 and PageMax > 1 then
			start_page = 2
			end_page = PageMax - 1
		elseif NowPage > 2 and NowPage <= PageMax -1 then
			start_page = NowPage - 1
	        end_page = start_page + 7
			if end_page > PageMax - 1 then
				end_page = PageMax - 1
				start_page = end_page - 7
			end if
		elseif NowPage = PageMax then
			end_page = NowPage - 1
			start_page = end_page - 7
		else
			start_page = 2
			end_page = start_page + 7
		end if

%>
<div class='row'>
    <div class='pagination'>
    <form action='<%=ActFile%>?<%=other_args%>' method=post>
        <ul>
            <li class='prev <%=prev_disabled%>'><a href='<%=prev_url%>'>&larr;<%=jm("__FIRST_PAGE__")%></a></li>
    <% if NowPage  = 1 then active = "active" else active = "" %>
            <li class='<%=active%>'><a href='<%=first_url%>'>1</a></li>
    <% if start_page > 2 then %>
            <li class=''><a href='#'>...</a></li>
    <% end if %>
    <% for p_i = start_page  to end_page
            if clng(NowPage) = p_i then active = "active" else active = ""
    %>
            <li class='<%=active%>'><a href='<%=ActFile & "?page=" & p_i & other_var %>'><%=p_i%></a></li>
    <% next %>
    <% if end_page < PageMax - 1 then %>
            <li class=''><a href='#'>...</a></li>
    <% end if %>
    <% if clng(NowPage)  = PageMax then active = "active" else active = "" %>
            <li class='<%=active%>'><a href='<%=last_url%>'><%=PageMax%></a></li>
            <li class='next <%=next_disabled%>'><a href='<%=next_url%>'>&rarr;<%=jm("__NEXT_PAGE__")%></a></li>
        </ul>
    </form>
    </div>
</div>
<%
	end sub

	public sub this_page ( str_ )
		page = mynumeric(str_)
		if page < 1 then page = 1
		
	end sub

	public property let sPkey ( str_ )
		pkey = str_
	end property

	public property let sPageSize ( str_ )
		page_size = str_
	end property

	Private Sub Class_Initialize 
		view = true
		del = false
		edit = false
		pkey = "sn"
		page_size = 20
		desc_ = "desc"
	End Sub

	Private Sub Class_Terminate
	End Sub

end class
%>
