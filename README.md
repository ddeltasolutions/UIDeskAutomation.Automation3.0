# UIDeskAutomation.Automation3.0
This is a .NET library that can be used to automate Windows desktop programs based on their user interface. It is created on top of unmanaged Microsoft UI Automation API (Windows Automation API 3.0)..

Here is an example. You can start an application like this:

<i><b>using UIDeskAutomationLib;<br>
...<br>
var engine = new Engine();<br>
engine.StartProcess("notepad.exe");<br></b></i>

After starting the application, you can set text to Notepad like this:

<i><b>UIDA_Window notepadWindow = engine.GetTopLevel("Untitled - Notepad");<br>
UIDA_Edit edit = notepadWindow.Edit();<br>
edit.SetText("Some text");<br></b></i>

You can find more information about this library <a href="http://automationspy.freecluster.eu/uideskautomation_unmanaged.html">HERE</a>.<br>
Install the library from <a href="https://www.nuget.org/packages/UIDeskAutomation.Automation3.0/">Nuget</a>.
