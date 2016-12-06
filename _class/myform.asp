<%
Class myform

	Private myclass_name 

	Public Sub set_class ( class_ )
		myclass_name = class_	
	end sub

	Private Sub Class_Initialize 

	End Sub

	public sub add_hidden ( name , value )
		response.write "<input type=hidden name='" & name & "' value='" & value & "'>"
	end sub

	public sub add_checkbox_s ( title , name , value , value2 , title2 , class_name , help )
		if value2 = value or value = true then checked = " checked " else checked = ""

		str_ = "<label class='checkbox inline'><input type=checkbox name='" & name & "' " & checked & " value='" & value2 & "'>" & title2 
		if len(help) > 0 then
			str_ = str_ & "<span class='help-inline'>" & help & "</span>"
		end if

		call show_base ( title , str_ )
	end sub

	public sub show_yes ( title , yes_) 
		if value = "Y" then yes_ = 1
		yes_ = mynumeric(yes_)
		if yes_ then
			y_img = "<img src='img/tick.png' border=0 alt='Yes'>"	
		else
			y_img = "<img src='img/cross.png' border=0 alt='No'>"	
		end if
		
		str_ = "<span class='inline'>" & y_img & "</span>"
		call show_base ( title , str_)
		
	end sub

	public sub add_radio ( title , class_name , name , value , items,  help)
		is2array = chk2array(items)

		str_ = ""
		if isnull(value) then value = ""

		if is2array then
			item_max = items(0,0)
			for item_i = 1 to item_max 
				if instr(cstr(value),cstr(items(item_i,1))) > 0 then checked = " checked " else checked = ""
				str_ = str_ & "<label class='radio inline'><input type=radio name='" & name & "' value='" & items(item_i,1) & "' " & checked & ">" & items(item_i,2) & "</label>" & vbcrlf
			next
		else
			for each item in items
				if instr(cstr(value),cstr(item)) > 0 then checked = " checked " else checked = ""
				str_ = str_ & "<label class='radio inline'><input type=radio name='" & name & "' value='" & item & "' " & checked & ">" & item & "</label>" & vbcrlf
			next
		end if

		if len(help) > 0 then
			str_ = str_ & "<span class='help-inline'>" & help & "</span>"
		end if

		call show_base ( title , str_ )

	end sub

	public sub add_checkbox ( title , class_name , name , value , items,  help)
		if isnull(value) then value = ""
		is2array = chk2array(items)

		str_ = ""

		if is2array then
			item_max = items(0,0)
			for item_i = 1 to item_max 
				if instr(cstr(value),cstr(items(item_i,1))) > 0 then checked = " checked " else checked = ""
				str_ = str_ & "<label class='checkbox inline'><input type=checkbox name='" & name & "' value='" & items(item_i,1) & "' " & checked & ">" & items(item_i,2) & "</label>" & vbcrlf
			next
		else
			for each item in items
				if instr(cstr(value),cstr(item)) > 0 then checked = " checked " else checked = ""
				str_ = str_ & "<label class='checkbox inline'><input type=checkbox name='" & name & "' value='" & item & "' " & checked & ">" & item & "</label>" & vbcrlf
			next
		end if

		if len(help) > 0 then
			str_ = str_ & "<span class='help-inline'>" & help & "</span>"
		end if

		call show_base ( title , str_ )

	end sub

	private sub show_base ( title , str_ )
%>
		<div class='control-group'>
			<label class='control-label'><%=title%></label>
			<div class='controls'>
				<%=str_%>
			</div>
		</div>
<%
	end sub

	
	public sub show_text ( title , class_name , value )

		str_ = "<input type=text class='" & class_name & "' disabled value='" & value & "'>"
		call show_base ( title , str_ )

	end sub

	public sub add_file2 ( title , class_name, name , value , tip , button_title)

		str_ = 	"<input type='file' name='" & name & "_file' id='" & name & "_file' class='" & class_name & "'>" & _ 
				"<p><button class='btn file_upload' value='" & name & "_file' id='" & name & "_button'>" & button_title & "</button>" & _
				"<div class='progress progress-info' id='" & name  & "_progress' style='display:none'><div class='bar' style='width: 60%'></div></div></p>" & _
				"<input type='hidden' name='" & name & "' value='" & value & "'>" & _	
				"<span class='help-block' id='" & name & "_tip'>" & tip & "</span>" 

		del_button = "<button class='btn btn-danger remove_button' data-id='" & name & "'>刪除檔案</button>"

		if chk_file(value) then 
			display = "display:block"
		else
			display = "display:none"
		end if
			
		str_ = str_ & "<div id='show_file_" & name & "' style='"  & display & "'><a href='" & value & "' target='download'>檔案檢視下載</a>&nbsp;&nbsp;" & del_button & "</div>" 


		call show_base ( title , str_ )

	end sub

	public sub add_file ( title , class_name, name , value , tip , button_title)

		str_ = 	"<input type='file' name='" & name & "_file' id='" & name & "_file' class='" & class_name & "' accept='image/*'>" & _ 
				"<p><button class='btn file_upload' value='" & name & "_file' id='" & name & "_button'>" & button_title & "</button>" & _
				"<div class='progress progress-info' id='" & name  & "_progress' style='display:none'><div class='bar' style='width: 60%'></div></div></p>" & _
				"<input type='hidden' name='" & name & "' value='" & value & "'>" & _	
				"<span class='help-block' id='" & name & "_tip'>" & tip & "</span>" 

		if chk_file(value) then 
			str_ = str_ & "<p><img src='" & value & "' class='img-polaroid' id='show_image_" & name & "'></p>" 
		else
			str_ = str_ & "<p><img src='' class='img-polaroid' id='show_image_" & name & "' style='display:none'></p>" 
		end if

		call show_base ( title , str_ )

	end sub

	public sub add_address ( title , class_name , zip_name , name , value )

		str_ = "<div id=" & zip_name & "></div><input type=text name=" & name & " class='" & class_name & "' value='" & value & "'>"
		call show_base ( title , str_ )

	end sub

	public sub show_help ( help  )
		str_ = "<span class='alert alert-info'>" & help & "</span>"
		call show_base ( title , str_ )
	end sub

	public sub add_textarea ( title , class_name , name , value )
		rows_t = clng(mynumeric(replace(class_name,"span",""))) * 3
		if rows_t = 0 then rows_t = 18
		str_ = "<textarea name='" & name & "' class='" & class_name & "' id='" & name & "' rows='" & rows_t & "'>" & value & "</textarea>"
		call show_base ( title , str_ )

	end sub

	Public Sub add_select ( title , class_name , name , value , items , empty_a , help) 
		if isnull(value) then value = ""

		is2array = chk2array(items)

		str_ = "<select name='" & name & "' class='" & class_name & "'>" & vbcrlf
		if isarray(empty_a) then
			if empty_a(0) then
				str_ = str_ &  "<option value='" & empty_a(1) & "'>" & empty_a(2) & "</option>" & vbcrlf
			end if
		end if

		if is2array then
			item_max = items(0,0)
			for item_i = 1 to item_max 
				if cstr(trim(items(item_i,1))) = cstr(trim(value)) then selected = " selected " else selected = ""
                if len(items(item_i,1)) > 0 then
    				str_ = str_ & "<option value='" & items(item_i,1) & "' " & selected & ">" & items(item_i,2) & "</option>" & vbcrlf	
                end if
			next
		else
			if isarray(items) then
				for each item in items
					if trim(item) = trim(value) then selected = " selected " else selected = ""
                    if len(item) > 0 then
			    		str_ = str_ & "<option value='" & item & "'"  & selected & ">" & item & "</option>" & vbcrlf
                    end if
				next
			end if
		end if

		str_ = str_ &  "</select>"
		if len(help) > 0 then
			str_ = str_ & "<span class='help-inline'>" & help & "</span>"
		end if
		call show_base ( title , str_ )

	END sub

	public sub add_type ( title , class_name , name , value , help , type_)
		str_ = "<input type='" & type_ & "' class='" & class_name & "' name='" & name & "' value='" & value & "'>"	
		if len(help) > 0 then
			str_ = str_ & "<span class='help-inline'>" & help & "</span>"
		end if
		call show_base ( title , str_ )
	end sub

	Public Sub add_text ( title , class_name , name , value , help)

		str_ = "<input type='text' class='" & class_name & "' name='" & name & "' value='" & value & "'>"	
		if len(help) > 0 then
			str_ = str_ & "<span class='help-inline'>" & help & "</span>"
		end if
		call show_base ( title , str_ )

	END sub

	Public Sub add_password ( title , class_name , name , value )

		str_ = "<input type='password' class='" & class_name & "' name='" & name & "' value='" & value & "'>"	
		call show_base ( title , str_ )

	END sub


	public sub add_progress ( id )
		str_ = "<div class='progress progress-striped active' id='" & id & "'><div class='bar' style='width:40%'></div></div>"
		call show_base ( "", str_)

	end sub
End Class
%>
