<div align="center">

## ASP and XML Surveys


</div>

### Description

Ever wanted a survey on your site? But don't have a database? Well, by using ASP and XML you can! Features single voting, survey expiry, past survey results, and many nore...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[George Leithead](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/george-leithead.md)
**Level**          |Intermediate
**User Rating**    |3.7 (26 globes from 7 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__4-7.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/george-leithead-asp-and-xml-surveys__4-6860/archive/master.zip)





### Source Code

		You've seen them all over the web. They've popped up more and more. What am I talking about? I'll tell you wat I'm talking about, <b>surveys</b> and user <b>poll</b>'s, thats what. Well, I have some code here that will allow you to implement a survey all without a database and other expensive components.<br />
		<br />
		All code mentioned in this article can be downloaded from <a href="http://forums.internetwideworld.com/CodeSamples/ASP_and_XML_Surveys.zip">here</a>, can can be discussed <a href="http://forums.internetwideworld.com/Default.asp?DefaultPage=ViewCompleteThread.asp%3FCategoryID%3D62%26ThreadID%3D48">here</a><br />
		<br />
		What do we need from a survey? Well, simply we need a question, answers to the question, and the ability to record the number of votes. This is acheived with the following simply XML structure:<br />
		<font color="darkgreen">
		&lt;survey voters=&quot;1&quot;&gt;<br />
		&nbsp;&nbsp;&nbsp;&nbsp;&lt;id /&gt;<br />
		&nbsp;&nbsp;&nbsp;&nbsp;&lt;question /&gt;<br />
		&nbsp;&nbsp;&nbsp;&nbsp;&lt;answer votes=&quot;1&quot; /&gt;<br />
		&nbsp;&nbsp;&nbsp;&nbsp;&lt;dateend /&gt;<br />
		&lt;/survey&gt;<br />
		</font>
		<br />
		Breaking this down, we can see that we have a survey node. We can have as many survey nodes as we like. This survey node has one attribute, <b>voters</b>. This will be used to record the number of votes taken for this survey. Each survey, has an ID, question and answer element(s). The id is used to uniquley identify each survey. Each survey has a question, and each survey can have one or more answers. After all a question with only one answer would be prity pointless. The last element is date end. This is used to specify when the survey is to end.<br />
		<br />
		You'll also notice that the answer element has a <b>votes</b> attribute. This will be used to record the number of votes for the selected answer.<br />
		<br />
		An exam of a populated survey might be:<br />
		<font color="darkgreen">
		&lt;survey voters=&quot;1&quot;&gt;<br />
		&lt;id /&gt;1&lt;/id&gt;<br />
		&lt;question&gt;Do you think surveys are a good idea?&lt;/question&gt;<br />
		&lt;answer votes=&quot;0&quot;&gt;Great&lt;/answer&gt;<br />
		&lt;answer votes=&quot;1&quot;&gt;Not Bad&lt;/answer&gt;<br />
		&lt;answer votes=&quot;0&quot;&gt;Average&lt;/answer&gt;<br />
		&lt;answer votes=&quot;0&quot;&gt;Poor&lt;/answer&gt;<br />
		&lt;answer votes=&quot;0&quot;&gt;Rubbish!&lt;/answer&gt;<br />
		&lt;dateend&gt;20011231&lt;/dateend&gt;<br />
		&lt;/survey&gt;<br />
		</font>
		<br />
		Now that we have defined the information we need, next we need to define a way to display it. By using XSL this is made relativley simple:<br />
		<font color="darkgreen">
&lt;render id=&quot;vradio&quot;&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;FORM ACTION=&quot;Poll.asp&quot; TARGET=&quot;Poll&quot; METHOD=&quot;POST&quot; onSubmit=&quot;window.open('','Poll','width=300,height=120,toolbar=no,location=no,directories=no,status=no,menubar=no,resizable=no,scrollbars=no');&quot;&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;INPUT TYPE=&quot;HIDDEN&quot; NAME=&quot;ID&quot;&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;xsl:attribute name=&quot;VALUE&quot;&gt;&lt;xsl:value-of select=&quot;id&quot;/&gt;&lt;/xsl:attribute&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;/INPUT&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;TABLE BORDER=&quot;1&quot; CELLPADDING=&quot;0&quot; CELLSPACING=&quot;0&quot; WIDTH=&quot;280&quot; SUMMARY=&quot;Survey&quot;&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TH&gt;Survey&lt;/TH&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TD ALIGN=&quot;center&quot;&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TABLE BORDER=&quot;0&quot; CELLPADDING=&quot;0&quot; CELLSPACING=&quot;0&quot; WIDTH=&quot;100%&quot; HEIGHT=&quot;100%&quot; SUMMARY=&quot;Survey Question and Answers&quot;&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TH COLSPAN=&quot;2&quot;&gt;&lt;xsl:value-of select=&quot;question&quot;/&gt;&lt;/TH&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;xsl:for-each select=&quot;answer&quot;&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TD &gt;&lt;xsl:value-of /&gt;&lt;/TD&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TD &gt;&lt;INPUT TYPE=&quot;radio&quot; NAME=&quot;Poll&quot;&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;xsl:attribute name=&quot;value&quot;&gt;&lt;xsl:value-of/&gt;&lt;/xsl:attribute&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/INPUT&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TD&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/xsl:for-each&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TD ALIGN=&quot;right&quot; COLSPAN=&quot;2&quot;&gt;&lt;INPUT TYPE=&quot;submit&quot; NAME=&quot;Submit&quot; VALUE=&quot;Vote&quot; /&gt;&lt;/TD&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TABLE&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TD&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TABLE&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;/FORM&gt;<br />
&lt;/render&gt;<br />
<br />
&lt;render id=&quot;results&quot;&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;TABLE BORDER=&quot;1&quot; CELLPADDING=&quot;0&quot; CELLSPACING=&quot;0&quot; WIDTH=&quot;280&quot; SUMMARY=&quot;Survey Results&quot;&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TH&gt;Survey Result&lt;/TH&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TD ALIGN=&quot;center&quot;&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TABLE BORDER=&quot;0&quot; CELLPADDING=&quot;0&quot; CELLSPACING=&quot;0&quot; WIDTH=&quot;100%&quot; HEIGHT=&quot;100%&quot; SUMMARY=&quot;Survey Result&quot;&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TH COLSPAN=&quot;3&quot;&gt;&lt;xsl:value-of select=&quot;question&quot;/&gt;&lt;/TH&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TD COLSPAN=&quot;3&quot; ALIGN=&quot;CENTER&quot;&gt;Total Voters: &lt;xsl:value-of select=&quot;@voters&quot;/&gt;&lt;/TD&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;xsl:for-each select=&quot;answer&quot;&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TD ALIGN=&quot;RIGHT&quot;&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;IMG SRC=&quot;Images/Vote.gif&quot; HEIGHT=&quot;10&quot; ALT=&quot;Votes&quot;&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;xsl:attribute name=&quot;WIDTH&quot;&gt;&lt;xsl:eval&gt;Math.round(this.getAttribute(&quot;votes&quot;)/this.parentNode.getAttribute(&quot;voters&quot;) * 100)&lt;/xsl:eval&gt;&lt;/xsl:attribute&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/IMG&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TD&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TD ALIGN=&quot;RIGHT&quot;&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;xsl:eval&gt;Math.round(this.getAttribute(&quot;votes&quot;)/this.parentNode.getAttribute(&quot;voters&quot;) * 100)&lt;/xsl:eval&gt;%<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TD&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TD WIDTH=&quot;100%&quot;&gt;&lt;xsl:value-of/&gt;&lt;/TD&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/xsl:for-each&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TABLE&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TD&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;TD ALIGN=&quot;center&quot;&gt;&lt;A HREF=&quot;ViewPastPolls.asp&quot;&gt;View Results of past surveys&lt;/A&gt;&lt;IMG SRC=&quot;http://forums/internetwideworld.com/Images/Blank.gif&quot; height=&quot;0&quot; width=&quot;0&quot; alt=&quot;Weaval&quot; /&gt;&lt;/TD&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TR&gt;<br />
&nbsp;&nbsp;&nbsp;&nbsp;&lt;/TABLE&gt;<br />
&lt;/render&gt;<br />
</font>
Now we have our information, and a template to display it in, we now need to select the correct survey for display. Firstly, we need to load the XML and XSL files:<br />
		<font color="darkgreen">
<pre>
Application("SurveyPath")  = server.MapPath("src/") & "\" ' NEED to add on the extra \
Application("SurveySource") = server.MapPath("src/surveys.xml")
Application("SurveyStyle") = server.MapPath("src/surveys.xsl")
Function loadSurveys()
 Dim source, style, rootPath
 rootPath = Application("SurveyPath")
 set source = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")
 source.async = false
 source.load Application("SurveySource")
 set style = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")
 style.async = false
 style.load Application("SurveyStyle")
 application.lock
 set application("survey") = source
 set application("style") = style
 application.unlock
end Function
</pre>
</font>
Next we need to actually display the survey. This is acheived by selecting the first survey where the dateend is greater than today, we can select the correct survey.<br />
		<font color="darkgreen">
<pre>
Function GetDate()
	'Produces YYYYMMDD
	DIM dYear : dYear = Year(Now())
	DIM dMonth : dMonth = Month(Now())
	DIM dDay  : dDay = Day(Now())
	DIM iID  : iID = 1
	DIM iLoop : iLoop = 1
	If Len(dMonth) < 2 then dMonth = "0" & dMonth
	If Len(dDay) < 2 then dDay = "0" & dDay
	GetDate = dYear & dMonth & dDay
End Function
Function DisplayLatestSurvey()
	DIM oRequiredNode
	DIM oSource
	DIM oStyle
	DIM oOutNode
	DIM sRender : sRender = "vradio"
	DIM iID  : iID = 1
	DIM dDate : dDate = GetDate()
	set oSource = Application("survey")
	set oStyle = Application("style")
	set oRequiredNode = oSource.selectSingleNode("surveylist/survey[dateend[. > """ & dDate & """]]")
	if NOT oRequiredNode is nothing then
		iID = oRequiredNode.selectSingleNode("id").text ' Get the Survey ID!
		If InStr(Request.Cookies("Survey"),":" & iID & ":") > 0 then
			sRender = "results"
		End If
		set oOutNode = oStyle.selectSingleNode("xsl:stylesheet/render[@id=""" & sRender & """]")
		if NOT oOutNode is nothing then
			' Apply the style to the source!
			DisplayLatestSurvey = oRequiredNode.transformNode(oOutNode)
		else
			DisplayLatestSurvey = "ERROR: Either the XSL file is incorrect or the render identifier is not found"
		end if
	else
		DisplayLatestSurvey = "ERROR: Either no survey exists for this period or the survey was not found"
	end if
	' Tidy Up!!!
	set oOutNode = nothing
	set oRequiredNode = nothing
	set oStyle = nothing
	set oSource = nothing
End Function
</pre>
</font>
Within the DisplayLatestSurvey function there are two items to pay particular attention to. Firstly, <b>sRender = "vradio"</b> specifies the render style we want, which is contained within the XSL file.<br />
The second item of note, is the <b>Request.Cookies</b>. Obviously, if some one has already voted in the survey, we don't want them to vote again. So, we determine if the user has a cookie this survey. If they do not, simply display the vradio XSL display. Oterhwise, they have already voted, and we change the XSL render to be <b>results</b>. This XSL render displays the actual voting for the survey.<br />
So far, we've only displayed the survey and the survey results. But what about letting the user register their vote? This is by adding 1 to not only the number of <b>voters</b>, but also by adding 1 to the number of <b>votes</b> for an answer.<br />
<font color="darkgreen">
<pre>
function voteSurvey(ByVal sID, ByVal sVote, ByRef bResult)
 Dim findNode, result, rootPath, source
 if sVote<>"" then
  set source = application("survey")
  set findNode = source.selectSingleNode("surveylist/survey[id[. = """ & sID & """]]/answer[.=""" & sVote & """]")
  if not findNode is nothing then
	  findNode.setAttribute "votes", findNode.getAttribute("votes") + 1
	  set findNode = findNode.parentNode
	  findNode.setAttribute "voters", findNode.getAttribute("voters") + 1
		 rootPath = Application("SurveyPath")
	  saveSurveys source
	  bResult = true
	  result = "Your vote has been accepted!"
  else
	 result="Your vote choice for this poll is invalid!"
  end if
 end if
 voteSurvey = result
 if bResult then CALL SetSurveyCookie(sID)
end function
function saveSurveys(obj)
 Dim rootPath,s
 rootPath = Application("SurveyPath")
 obj.save rootPath & "surveys.xml"
 for each s in application.contents
	if left(s,6)="SURVEY" and s<>"SURVEY" then
		application.lock
		set application(s) = nothing
		application.unlock
	end if
 next
end function
Function SetSurveyCookie(ByVal iID)
	Response.Cookies("Survey") = ":" & iID & ":"
	Response.Cookies("Survey").Expires = "January 1, 2011"
End Function
</pre>
</font>
Once we register the vote, we need to then save the results, and 'unload' the previous version of the XML file, so that the new results can be viewed.<br />
Then in order to prevent the user from re-submitting another vote, we send the user a cookie with the survey id that they voted in.<br />
<br />
As additional functionality, we may also want to display a list of 'past' surveys and their results. This is acheived by implementing a new XSL render template, and some more ASP code. These and more can be found in the downloadable code, from <a href="http://forums.internetwideworld.com/CodeSamples/ASP_and_XML_Surveys.zip">here</a>.<br />
<br />
To discuss this article and the code go <a href="http://forums.internetwideworld.com/Default.asp?DefaultPage=ViewCompleteThread.asp%3FCategoryID%3D62%26ThreadID%3D48">here</a>.<img src="http://forums.internetwideworld.com/Images/Blank.gif" width="0" height="0" alt="Weaval"/>

