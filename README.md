<div align="center">

## Free Threading in VB\.NET


</div>

### Description

An explanation of threading in VB.NET. This new exciting feature lets you FINALY execute code in an async fashion.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jeroen Schouten](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeroen-schouten.md)
**Level**          |Advanced
**User Rating**    |4.8 (124 globes from 26 users)
**Compatibility**  |VB\.NET
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__10-1.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeroen-schouten-free-threading-in-vb-net__10-658/archive/master.zip)





### Source Code

```
<style>
<!--
 /* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
p.Code, li.Code, div.Code
	{mso-style-name:Code;
	mso-style-next:Normal;
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:"Courier New";
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-family:"Times New Roman";}
@page Section1
	{size:8.5in 11.0in;
	margin:1.0in 1.25in 1.0in 1.25in;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
-->
</style>
</head>
<body lang=EN-US style='tab-interval:.5in'>
<div class=Section1>
<p class=MsoNormal>In this article I will try and give you an example of one of
the most exciting new feature of VB.NET… Free threading! Take a look the
following VB6 code:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=Code>Dim myTest as Ctest</p>
<p class=Code><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=Code>Public Function TestFunction(Id as Integer) as Integer</p>
<p class=Code><span style='mso-tab-count:1'>      </span>Dim Ret as Integer</p>
<p class=Code>Set myTest = New<span style="mso-spacerun: yes">  </span>Ctest</p>
<p class=Code><span style='mso-tab-count:1'>      </span>Ret =
MyTest.DoSomething(Id)</p>
<p class=Code><span style='mso-tab-count:1'>      </span>TestFunction = Ret</p>
<p class=Code><span style='mso-tab-count:1'>      </span>Set myTest = Nothing</p>
<p class=Code>End Function</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>What happens here is very simple. We have a class (Ctest)
which contains a method called: DoSomething. The DoSomething method takes one
argument (Id) and returns an integer. The above piece of code wraps all this
neatly in a function. I am sure you’ve done something like this a million times
before.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Now what happens when DoSomething is a method that runs for
a very long time? Let’s say it iterates through an ADO Recordset containing a
thousand records? Right! Your code is going to hit the line:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=Code><span style='mso-bidi-font-family:"Courier New"'>Ret =
MyTest.DoSomething(Id)<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>…And wait till the DoSomething method is done before your
code will continue with the line below it (the one that sets TestFunction to
the return value of DoSomething). While this is acceptable in most cases, for
instance if DoSomething was to contain the code that saves a long document and
the user has to wait anyway, it is very annoying behavior in a lot of other
cases. I have tried every trick in the book, but VB6 doesn’t have a neat way of
handling this situation. At some point your code will be waiting for other code
to finish. If you ever tried keeping the GUI responsive with DoEvents while
your application is off to do some expensive database operation, you will know
exactly what I mean.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>VB.NET to the rescue! Enter stage left, FREE THREADING.
Wouldn’t it be great if you had a way to tell VB that you really don’t need to
wait but instead want to go on with whatever the next task of your code is? You
can!</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>The trick is in: </p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'>System.Threading.Thread<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal>The system.threading namespace allows you to push methods of
a class or subs and functions of your regular code off to its own thread.
Create a form with a button on it called Button1. Then…</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=Code>‘the following goes under the Inherits portion of your VB code</p>
<p class=Code>Dim Test as New Ctest</p>
<p class=Code>Dim ThreadTest as System.Threading.Thread(AddressOff
MyThreadTest)</p>
<p class=Code><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=Code>‘this is a simple Sub:</p>
<p class=Code>Public Sub MyThreadTest</p>
<p class=Code><span style='mso-tab-count:1'>      </span>Test.DoSomeThing(Id)</p>
<p class=Code>End Sub</p>
<p class=Code><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=Code>‘this is the button’s click event</p>
<p class=Code><span style='mso-bidi-font-size:10.0pt;mso-bidi-font-family:"Courier New"'>Private
Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
Handles Button1.Click<o:p></o:p></span></p>
<p class=Code style='text-indent:.5in'><span style='mso-bidi-font-size:10.0pt;
mso-bidi-font-family:"Courier New"'>MsgBox(&quot;Beginning New Threads&quot;)<o:p></o:p></span></p>
<p class=Code style='margin-left:.5in'><span style='mso-bidi-font-size:10.0pt;
mso-bidi-font-family:"Courier New"'>ThreadTest = New
System.Threading.Thread(AddressOf _ MyThreadTest)<o:p></o:p></span></p>
<p class=Code style='text-indent:.5in'><span style='mso-bidi-font-size:10.0pt;
mso-bidi-font-family:"Courier New"'>ThreadTest.Start()<o:p></o:p></span></p>
<p class=Code><span style='mso-tab-count:1'>      </span>MsgBox(“Done”)</p>
<p class=Code><span style='mso-bidi-font-size:10.0pt;mso-bidi-font-family:"Courier New"'>End
Sub<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Do NOT run this code! IT WILL FAIL. There is a couple of
things we still need to do, but the above part shows us what is going on when
we get into multi-threading. So what the heck is going on??</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>The first line we simply set a reference (Dim) to our Ctest
class (remember? The one with the very loooooong running method.) In the next
line we simply tell VB to declare a new thread! That is really most of what
there is to it. By saying “Dim &lt;somename&gt; as new
System.Threading.Thread(AddressOf &lt;somemethod&gt;)” We have created a new
thread to call a method/sub on. The AddressOf operator might be a surprise to
you. What you are doing there is saying to VB “Hey when I start doing something
on that new thread I want you to be doing &lt;insert name of method/sub&gt;”.
Some people might now say that I am over simplifying this… I AM! But for now
that is all you need to know.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>You can see that I pointed the thread to a Sub
(AddressOf<span style="mso-spacerun: yes">  </span>MyThreadTest ‘MyThreadTest
is a Sub) and not directly to the class we are using. There is a very good
reason for that. If you want anything to run in its own thread you can’t pass
arguments, or receive return values. I know what you are thingking…. “Geez,
that’s not very useful!” But think about it for a second. You are telling VB to
start a new thread… A NEW THREAD! Not a new instance of a class, not another
copy of an ActiveX Server… A NEW THREAD! You completely disconnected the code
on that thread from the main application. The application you’re doing this in
doesn’t know anything about this new thread, and quite frankly, it doesn’t give
a rodent’s butt-cheek. That is why you we need to wrap that code into a sub
that our application DOES know about and that is exactly what I am doing this
piece of code:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=Code>Public Sub MyThreadTest</p>
<p class=Code><span style='mso-tab-count:1'>      </span>Test.DoSomeThing(Id)</p>
<p class=Code>End Sub</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>I set my new thread to be: AddressOf MyThreadTest. Even
though The Sub doesn’t take any arguments or returns anything, I can have the
call to the Ctest class inside of it take an argument (Id). Oh! This is why I
said not to run it yet previously. VB.NET will complain that it doesn’t know “Id”.
For testing purposes we can quickly fix this. Change the code of the Sub
MyThreadTest to read:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=Code>Public Sub MyThreadTest</p>
<p class=Code><span style='mso-tab-count:1'>      </span>Dim Id as Integer</p>
<p class=Code><span style='mso-tab-count:1'>      </span>Id = 10</p>
<p class=Code><span style='mso-tab-count:1'>      </span>Test.DoSomeThing(Id)</p>
<p class=Code>End Sub</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Of course to run all this you better create a class in your
project called Ctest with a DoSomething Function that takes an Integer (Id) as
an argument, but you get the idea.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>The last code in your project is the Click event of the
button I had you put on the form. The most important line of code inside it is:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=Code>ThreadTest.Start()</p>
<p class=MsoNormal><span style='mso-bidi-font-size:10.0pt;mso-bidi-font-family:
"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='mso-bidi-font-size:10.0pt;mso-bidi-font-family:
"Courier New"'>See what we are doing here is not calling the Sub MyThreadTest.
Instead we are saying: I declared a thread called ThreadTest, Start it! Because
we previously set the ThreadTest to be AddressOf MyThreadTest, VB will now run
the code inside that Sub in its own thread.<o:p></o:p></span></p>
<p class=MsoNormal><span style='mso-bidi-font-size:10.0pt;mso-bidi-font-family:
"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='mso-bidi-font-size:10.0pt;mso-bidi-font-family:
"Courier New"'>To get the full effect of this you should really add a class
called Ctest to your project and create a DoSomething function in it that does
something that takes very very very looooong. If you’ve done that, start the
application and click on the button. What you will see is the first message box
reading: Beginning New Threads. If you on OK, the code will jump into the Sub
MyThreadTest. But rather then waiting for execution of this Sub (and the call
to the Ctest method DoSomething inside it) it will almost immediately show the
next message box (“Done!”). And THAT is free threading in a nutshell… Cool or
what??</span></p>
</div>
```

