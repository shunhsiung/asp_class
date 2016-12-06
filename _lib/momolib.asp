<%
' Author: Eric Shih 
' Date: 2009.06.04

USER_IP = Request.ServerVariables("REMOTE_ADDR")

function chkregexp ( str1 , reg_str )
	set reg = new regexp
	with reg
		.pattern = reg_str
		.IgnoreCase = true
	end with

	chkregexp = reg.test(str1)
end function

sub println( str_ )
    response.write str_ & "<br>" & vbcrlf
end sub

sub prints( str_ )
    response.write str_ & vbcrlf
end sub

sub dprint(str_)
    select case USER_IP
      case "60.251.149.19","220.132.221.229"
        response.write "<font color=red>[debug]</font> " & str_ & "<br>" & vbcrlf
    end select
end sub

function mynumeric(t)
	mynumeric = t
    if len(t) = 0 or not isnumeric(t) then mynumeric = 0 
end function

function mydate( t )
   mydate = t
   if len(t) = 0 or not isdate(t) then mydate = date()
end function

function showToday_S 
    Today_S = date()      
    weekly = DatePart("w",Today_S) 
    select case weekly
       case "1"
          weekly_s = "天"
       case "2"
          weekly_s = "一"
       case "3"
          weekly_s = "二" 
       case "4"
          weekly_s = "三"
       case "5"
          weekly_s = "四"
       case "6"
          weekly_s = "五"
       case "7"
          weekly_s = "六"
    end select
    showToday_S = Today_S & " 星期" & weekly_s
end function

function chkMobile ( str_ )
	if isnull(str_) then
		chkMobile = false
		exit function
	end if

	set Reg = new RegExp
	with Reg
		.pattern = "^09\d{8}$"
		.IgnoreCase = True
	end with
	chkMobile = Reg.Test(str_)
end function

function chkUrl (str_)
    set Reg = new RegExp
    with Reg
	 .pattern = "^(http://|https://){0,1}[A-Za-z0-9][A-Za-z0-9\-\.]+[A-Za-z0-9]\.[A-Za-z]{2,}[\43-\176]*$"
     .IgnoreCase = True
    end with
    chkUrl = Reg.Test(str_)
end function

function chkEmail ( str_ )
    set Reg = new RegExp 
    with Reg
     .pattern = "^[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
     .IgnoreCase = True
    end with
    chkEmail = Reg.Test(str_)
end function

function chkChinese ( str_ )
    set Reg = new RegExp
    with Reg
      .pattern = "^[a-zA-Z0-9]*[a-zA-Z0-9]$"
      .IgnoreCase = False
    end with
    chkChinese = Not Reg.Test(str_)
end function

function chkPersonID ( id_ )
   if len(id_) <> 10 then chkPersonID = False : exit function
   set Reg = new RegExp
   with Reg
     .pattern = "^[a-z][1-2]\d{8}" 
     .IgnoreCase = True
   end with

   if Reg.Test( id_ ) then
     str1 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
     str2 = "1011121314151617341819202122352324252627282932303133"
     t1 = mid(str2,instr(str1,ucase(mid(id_,1,1)))*2-1, 2)
     
     sum = int(mid(t1,1,1)) * 1 + int(mid(t1,2,1)) * 9 
     
     t10 = mid(id_,10,1)

     for id_i = 2 to 9 
        sum = sum + int(mid(id_,id_i,1)) * (10 - id_i) 
     next
     
     if (sum mod 10) = 0 then
       t10_ = 0 
     else
       t10_ = 10 - (sum mod 10)
     end if

     if cint(t10_) = cint(t10) then 
        chkPersonID = True
     else 
        chkPersonID = False
     end if
   else
      chkPersonID = False
   end if

end function 

function Rspecial_str( str_ )
  ' &#34; => " &#38; => & &#39; => ` &#60; => < &#62; => > &#92; => \ &#96; => '
  str_ = replace(str_,chr(34),"&#34")
  str_ = replace(str_,chr(38),"&#38")
  str_ = replace(str_,chr(39),"&#39")
  str_ = replace(str_,chr(60),"&#60")
  str_ = replace(str_,chr(62),"&#62")
  str_ = replace(str_,chr(92),"&#92")
  str_ = replace(str_,chr(96),"&#96")
  Rspecial_str = str_
end function

function RHtmlTag( str_)
  set Reg = new RegExp 
  with Reg
   .pattern = "<.[^<]*>|</[^<]*>"
   .IgnoreCase = True
   .Global = True
  end with
  RHtmlTag = Reg.Replace(str_,"")
end function

sub mail_user ( from_ , to_ , subject_ , body_ , html) 

 Set newobjectmail = Server.CreateObject ("CDO.Message")
 
 with newobjectmail 
   .from = from_
   .to = to_ 
   .subject = subject_
   .BodyPart.Charset = "utf-8" 
   if html then
     .htmlbody = body_
   else
     .textbody = body_
   end if
   .send
 end with

end sub

function mv_file (file1,file2)
  dim fs
  set fs = server.createobject("scripting.filesystemobject")
  fs.MoveFile file1,file2
  set fs=nothing
end function 

function save_utf8_file (filename, content)
	fn = server.mappath(filename)  
	set txt = createobject("ADODB.Stream")
    txt.open
    txt.charset="utf-8"
    txt.writetext content
    txt.savetofile fn,2
	txt.close
  	set txt = nothing
end function

function save_file (filename, content)
  set fs = server.createobject("scripting.filesystemobject")
  filename = server.mappath(filename)  
  set fname = fs.CreateTextFile(filename,true)
  fname.write content
  fname.close
  set fname = nothing
  set fs = nothing
end function

function save_inc_file (filename, content)
  set fs = server.createobject("scripting.filesystemobject")
  set fname = fs.CreateTextFile(filename,true)
  fname.write content
  fname.close
  set fname = nothing
  set fs = nothing
end function

function del_file(file1)
  dim fs
  set fs = server.createobject("scripting.filesystemobject")
  if fs.FileExists(file1) then
     fs.deletefile(file1) 
  end if
  set fs=nothing
end function

function chk_file (file1)
	if len(file1) = 0 then 
		chk_file = false
		exit function
	end if
	dim fs
	set fs = server.createobject("scripting.filesystemobject")
    filename = server.mappath(file1)  
	chk_file = fs.FileExists(filename) 
	set fs = nothing
end function

function CreateThumbImage ( ImageF1 , ImageF2 , Width, Height)
  Dim MyObj
  Set MyObj = Server.CreateObject("GflAx.GflAx")
  f1 = server.mappath(ImageF1)
  f2 = server.mappath(ImageF2)

  MyObj.LoadBitmap f1

  if (MyObj.Width > Width) then
    Height_x = int((Width / MyObj.Width) * MyObj.Height )
    MyObj.Resize Width,Height_x
  end if

  MyObj.SaveBitmap f2

end function

sub show_cookies ( c_name )
    prints request.cookies(c_name)
end sub

function get_cookies_value ( c_name )
    get_cookies_value = request.cookies(c_name)
end function

sub set_cookies_value ( c_name , value_, days)
    response.cookies(c_name) = value_
    response.cookies(c_name).expires = dateadd("d",days,Date()) & " 23:59"
end sub

sub empty_cookies ( c_name )
    response.cookies(c_name) = empty
end sub

function conver_ftime ( day_ )
	conver_ftime = year(day_) & string(2-len(month(day_)),"0") & month (day_) & string(2-len(day(day_)),"0") & day(day_)

end function

sub mail_user_big5 ( from_ , to_ , subject_ , body_ , html)

 Set newobjectmail = Server.CreateObject ("CDO.Message")

 with newobjectmail
   .from = from_
   .to = to_
   .subject = subject_
   .BodyPart.Charset = "big5"
   if html then
     .htmlbody = body_
   else
     .textbody = body_
   end if
   .send
 end with

end sub


function getmyfname
   fname = split(Request.ServerVariables("SCRIPT_NAME") ,"/")
   getmyfname = fname(ubound(fname))
end function

function file2str ( filename )
    filename = server.mappath(filename)
    set asm = server.createobject("Adodb.Stream")
    asm.open
    asm.charset = "utf-8"
    asm.LoadFromFile filename
    str_t = asm.readtext
    asm.close

    file2str = str_t
end function


%>
