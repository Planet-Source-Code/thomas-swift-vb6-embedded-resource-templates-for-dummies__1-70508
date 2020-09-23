<div align="center">

## VB6 Embedded Resource Templates For Dummies


</div>

### Description

An explanation of windows resource templates.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2008-05-09 22:52:38
**By**             |[Thomas Swift](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/thomas-swift.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[VB6\_Embedd2112375102008\.zip](https://github.com/Planet-Source-Code/thomas-swift-vb6-embedded-resource-templates-for-dummies__1-70508/archive/master.zip)





### Source Code

<div align="center">
 <center>
<table border="0" width="652" cellspacing="0" cellpadding="0">
 <tr>
  <td width="650">
   <div align="center">
    <center>
    <table border="0" width="94%" cellspacing="0" cellpadding="0">
     <tr>
      <td width="100%">
       <p align="center"><font face="Arial" size="5"><b>VB6 Embedded Resource
   Templates For </b></font><font face="Arial" size="5"><b>Dummies</b></font></td>
     </tr>
     <tr>
      <td width="100%">
       <p align="center">
   <font face="Arial" size="2"><b>Ever look at Microsoft's Picture and Fax
   viewer control and wonder just how they pulled it off ?<br>
   Well then your
   reading the right document !</b></font></td>
     </tr>
    </table>
    </center>
   </div>
   <hr>
  </td>
 </tr>
 <tr>
  <td width="650"><font face="Arial" size="2"><b>First off their are a lot of
   ways you can pull off a image presentation from the depths of source code.
   The hardest way is to do it programicly considering how hard it is to binary
   decode files. Especially gif's. When you look at a directory thru windows
   file explorer all your doing is viewing it thru a template driven and
   scripted web browser control. In fact even your desktop is a heavily
   scripted web browser control believe it or not. If you don't believe me
   just point eather the VB web browser control or a sub classed copy of
   internet explorer at a file directory and you will find out what I mean.
   Even the explorer right click menu comes into view on right click and you
   can have iconic view.</b></font>
   <p><font face="Arial" size="2"><b>A closer look at the file that houses
   the Microsoft Picture viewer using Resource Hacker and behold what do you
   find but a HTML template with insert code here variables all over the
   place. They pulled the same thing off in their Web Photo Album Generator
   on the Windows 98 CD. In fact if you look closely at a lot of the Windows
   operating system you will find embedded templates being used all over the
   place. Some like the Web Photo Album Generator even use embedded cascading
   style sheets or CSS for short.</b></font></p>
   <p><font face="Arial" size="2"><b>Well I promised to teach you just how
   they are pulling this off. Well in the VB6 GUI you have the VB Resource
   Editor on your toolbar. In the resource editor you have 5 types of
   resources you can add to an embedded resource file. The one furthest to
   the right is Custom. Using the custom resource type you can add literally
   anything to your project. Once you have chosen and added your particular
   resource it assigns each one an index number. Numbers 101, 102, 103 on to infinity.
   This is how you store your HTML templates in your final program.</b></font></p>
   <p><font face="Arial" size="2"><b>Now that you know how to store a
   template you need to know how to get it out and view it. Well lets look at
   some simple code to extract a custom string resource and view it.</b></font></p>
   <p><textarea rows="24" cols="78" name="Code">Input: LoadQuickFeed 101, "http://www.sfgate.com/rss/feeds/bondage.xml"
Public Function LoadQuickFeed(DataIndex As Integer, FeedURL As String)
Dim myArray As String
Dim TheFile As Integer
WebBrowser1.Visible = True
myArray = ByteArrayToString(LoadResData(DataIndex, "CUSTOM"))
myArray = Replace(myArray, "<<--URL-HERE-->>", FeedURL)
TheFile = FreeFile()
Open App.Path & "\temp.htm" For Output As #TheFile
Print #TheFile, myArray
Close #TheFile
WebBrowser1.Navigate2 App.Path & "\temp.htm"
End Function
Private Function ByteArrayToString(bytArray() As Byte) As String
Dim sAns As String
Dim iPos As String
sAns = StrConv(bytArray, vbUnicode)
iPos = InStr(sAns, Chr(0))
If iPos > 0 Then sAns = Left(sAns, iPos - 1)
ByteArrayToString = sAns
End Function
</textarea></p>
   <p><font face="Arial" size="2"><b>In the code above I have a previously
   coded template embedded in my project as resource 101 written in Java
   Script AJAX to display the RSS feed URL passed in the FeedURL variable.
   The template has the </b></font><font face="Arial"><b>&lt;&lt;--URL-HERE--&gt;&gt;
   <font size="2">written in the place of the feed URL in it. I did that by
   editing the HTML file with a text editor like Note Tab Pro after creating
   it in Front Page and the Windows Scripting Editor.<br>
   The code simply replaces the </font>&lt;&lt;--URL-HERE--&gt;&gt;
   <font size="2">with what ever feed URL I want to view when its loaded.
   Presto I have a totally versatile way of loading RSS feeds.</font></b></font></p>
   <p><font size="2" face="Arial"><b>Well now that you know the secret the variables
   are should we say the sky is the limit ! Have fun expanding on this basic
   shell concept created by Microsoft eons ago ! Enjoy !!!</b></font><p><font size="2" face="Arial"><b>The
   accompanying download contains same tutorial, but includes the HTML
   template source example as well.</b></font></td>
 </tr>
</table>
 </center>
</div>

