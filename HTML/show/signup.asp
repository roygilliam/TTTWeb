<%
	if Request("submit") = "submit" then
	
		' Have all fields been chosen?
		strFields = ""
		if Trim(Request("txtName")) = "" or Trim(Request("txtEmail")) = "" or Trim(Request("txtGradYear")) = "" or Request("optGender") = "" or (Request("cboSat") = "" and Request("cboSun") = "") Then
		   
			strFields = "*** Please make sure to enter information all required information (Name, Email, Graduation Year and Time Slot). ***"
			
		Else
			If Request("cboSat") <> "" and Request("cboSun") <> "" Then
			
				strFields = "*** Please chose only one audition time slot. ***"
				
			End If
			
		End If
			
		if strFields = "" Then
	
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'' Customize the following 5 lines with your own information. ''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			vtoaddress = "roy.gilliam@gmail.com, ddeweese0245@wowway.com" ' Change this to the email address you will be receiving your notices.
			vmailhost = "mymail.brinkster.com"  ' Change this to mail.yourDomain or leave as is.
			vfromaddress = "auditions@tetelestaithemusical.com" ' Change this to the email address you will use to send and authenticate with.
			vfrompwd = "t3t3l3stai" ' Change this to the above email addresses password.
			vsubject = "2012 Tetelestai Audition Sign Up" 'Change this to your own email message subject.
			 
			'''''''''''''''''''''''''''''''''''''''''''
			'' DO NOT CHANGE ANYTHING PAST THIS LINE ''
			'''''''''''''''''''''''''''''''''''''''''''
			vmsgbody = "<b>2012 Tetelestai Audition Request</b><br><hr>" 
			vmsgbody = vmsgbody & "<br>"
			vmsgbody = vmsgbody & "Name: " & request("txtName") & "<br>"
			vmsgbody = vmsgbody & "<br>"
			vmsgbody = vmsgbody & "Email: " & request("txtEmail") & "<br>"
			vmsgbody = vmsgbody & "<br>"
			vmsgbody = vmsgbody & "Grad Year: " & request("txtGradYear") & "<br>"
			vmsgbody = vmsgbody & "<br>"
			vmsgbody = vmsgbody & "Gender: " & request("optGender") & "<br>"
			vmsgbody = vmsgbody & "<br>"
			vmsgbody = vmsgbody & "Time Slot: " 
			If Trim(request("cboSat")) <> "" Then
				vmsgbody = vmsgbody & "Saturday January 7 at " & request("cboSat") & "<br>"
			Else
				vmsgbody = vmsgbody & "Sunday January 8 at "& request("cboSun") & "<br>"
			End If
			vmsgbody = vmsgbody & "<br>"
			vmsgbody = vmsgbody & "Comments/Notes: " & request("txtComments") & "<br>"
			 
			Set objEmail = Server.CreateObject("Persits.MailSender")
			 
			objEmail.Username = vfromaddress
			objEmail.Password = vfrompwd
			objEmail.Host = vmailhost
			objEmail.From = vfromaddress
			objEmail.AddAddress vtoaddress
			objEmail.Subject = vsubject
			objEmail.Body = vmsgbody
			objEmail.IsHTML = True
			objEmail.Send

			Set objEmail = Nothing
			 
			vErr = Err.Description
			if vErr <> "" then
				strFields = "*** There was an error submitting your request.  Please try again.  If you continue to have problems please email "
				strFields = strFields & "<a href=""mailto: auditions@tetelestaithemusical.com"">auditions@tetelestaithemusical.com</a>. ***"
			else
				response.redirect("signup_success.asp?satslot=" & request("cboSat") & "&sunslot=" & request("cboSun"))
			End If
		   
		 End If
	
	end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"> 
<html xmlns="http://www.w3.org/1999/xhtml"> 
<head> 
<meta http-equiv="content-type" content="text/html; charset=utf-8" /> 
<title>Tetelestai The Musical - 2012 Audition Sign Up - Columbus, OH</title> 
<meta name="description" content="Sign up for te 2012 season in Columbus, OH." /> 
<meta name="keywords" content="Tetelestai the Musical It Is Finished Russ Nagy Joel Nagy UALC Upper Arlington Lutheran Church " /> 
<link href="/default.css" rel="stylesheet" type="text/css" media="screen" /> 
</head> 
<body> 
 
<div id="header"> 
	<div id="logo"> 
		<!--
		<h1><a href="#"><span>tet&#233;lestai</span></a></h1>
		<p>It Is Finished</p>
		--> 
	</div> 
	<div id="menu"> 
		<ul id="main"> 
			<li><a href="/default.htm">Home</a></li> 
			<li><a href="/show/default.htm">Shows & Tickets</a></li> 
			<li><a href="/information/default.htm">Information</a></li> 
			<li><a href="/alumni/default.htm">Alumni</a></li> 
			<li><a href="/history/default.htm">History</a></li> 
			<li><a href="/contact.htm">Contact Us</a></li> 
			<li><a href="/contact.htm">Donate</a></li> 
		</ul> 
		<!--
		<ul id="feed">
			<li><a href="#">Entries RSS</a></li>
			<li><a href="#">Comments RSS</a></li>
		</ul>
		--> 
	</div> 
	
</div> 
 
		
<div id="wrapper"> 
	<!-- start page --> 
	<div id="page"> 
 
		<!-- start content --> 
		<div id="content_wide"> 
			<div class="breadcrumb"> 
				<a href="/default.htm">Home</a> &gt;&gt; 
				2012 Season Audition Sign Up
			</div> 
			<!--
			<div class="post"> 
				<h1 class="title-no-byline"><a href="/show/faq.htm">2010 Season FAQ</a></h1> 
				<div class="entry"> 
					<p>Tetelestai FAQ posted with information on auditions, touring, show schedule, costs, etc.</p> 
				</div> 
			</div> 
			-->
			<div class="post"> 
				<h1 class="title-no-byline">2012 Audition Sign Up</h1> 
				<div class="entry"> 
					<p>
						Auditions are being held on Saturday January 7, 2012 from 10:00am to 12:00pm and on Sunday January 
						8, 2012 from 12:00pm to 3:00pm.  In order to guarantee a prompt audition experience for you we ask that 
						you sign up for a time slot below.  We will be auditioning 9 people each hour.
					</p>
<%
	If strFields <> "" Then
		Response.Write "<p><font color=red>" & strFields & "</font></p>"
	End If
%>
					<p>
<h1><font color="red">Please send an email to DeEtte DeWeese to schedule an audition: <a href="mailto: deettemail@gmail.com">deettemail@gmail.com</a></font></h1>
<!--
						<form action="signup.asp" method="post">
						
							<table>
								<tr>
									<td>Your Name:</td>
									<td>
										<input type="text" name="txtName" value="<%=Request("txtName")%>">
									</td>
									<td></td>
								</tr>
								<tr>
									<td nowrap>Your Email Address:</td>
									<td>
										<input type="text" name="txtEmail" value="<%=Request("txtEmail")%>">
									</td>
									<td></td>
								</tr>
								<tr>
									<td>Graduation Year:</td>
									<td>
										<input type="text" name="txtGradYear" value="<%=Request("txtGradYear")%>">
									</td>
									<td>
										<i>What year will you be graduating from high school?</i>
									</td>
								</tr>
								<tr>
									<td>Gender:</td>
									<td>
										<input type="radio" name="optGender" value="female" /> Female<br>
										<input type="radio" name="optGender" value="male" /> Male
									</td>
									<td>
										<i>This is optional to help us understand the mix of male/female voices we will have.</i>
									</td>
								</tr>
								<tr><td colspan="3">&nbsp;</td></tr>
								<tr>
									<td>Time Slot:</td>
									<td colspan="2">
										<i>The time slots are for an hour even though your actual audition will be much shorter than that.  An email will be sent to you prior to your audition day with a more specific time on when to show up within your hour, what to expect, how to prepare, etc.</i>
									</td>
								</tr>
								<tr>
									<td></td>
									<td colspan="2">
										Saturday January 7: 
										<select name="cboSat">
										  <option value=""></option>
										  <option value="10:00am">10:00am - 11:00am (1 slot available)</option>
										</select>
										<br>
										Sunday January 8: 
										<select name="cboSun">
										  <option value=""></option>
										  <option value="12:00pm">12:00pm - 1:00pm (3 slots available)</option>
										</select>
									</td>
									<td></td>
								</tr>
								<tr><td colspan="3">&nbsp;</td></tr>
								<tr>
									<td>Notes/Comments:</td>
									<td colspan="2">
										<textarea rows="5" cols="30" name="txtComments"></textarea>
									</td>
								</tr>
							</table>
							<input type="submit" name="Submit" value="submit">
						</form>
-->
					</p> 
				</div> 
			</div> 
		</div> 
		<!-- end content --> 
 

		<div style="clear: both;">&nbsp;</div> 
	</div> 
	<!-- end page --> 
</div> 
 
<div id="footer"> 
	<p class="copyright">&copy;&nbsp;&nbsp;2011 Tetelestai, Inc. All Rights Reserved &nbsp;&bull;&nbsp; Design by <a href="http://www.freecsstemplates.org/">Free CSS Templates</a>.</p> 
	<p class="link"><a href="/SiteMap.htm">Site Map</a><!--&nbsp;&#8226;&nbsp;<a href="#">Terms of Use</a>--></p> 
</div> 
<script type="text/javascript"> 
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script> 
<script type="text/javascript"> 
try {
var pageTracker = _gat._getTracker("UA-11812086-1");
pageTracker._trackPageview();
} catch(err) {}</script> 
 
</body> 
</html>