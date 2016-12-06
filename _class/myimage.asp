<%
Class myimage

	public sub thumbnails ( class_name , img_a )
%>
	<div class='controls'>
	<ul class='thumbnails'>
<%
	if isarray(img_a) then
		img_max = img_a(0,0) 	
		for m_i = 1 to img_max
%>
		<li class='<%=class_name%>' id='img_<%=img_a(m_i,1)%>'> 
			<div class='thumbnail'>
<%
			if len(img_a(m_i,2)) > 4 then
%>
			<img src='<%=img_a(m_i,2)%>' alt=''>
<% 			
			else 
%>
			<span class='label'>NoImage</span>
<%
			end if
%>
			</div>
			<h5><%=img_a(m_i,3)%></h5>
		</li>
<%
		next
	end if
%>
	</ul>
	</div>
<%	
	end sub
end Class
%>
