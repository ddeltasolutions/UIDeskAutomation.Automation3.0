# UIDeskAutomation.Automation3.0
This is a .NET library that can be used to automate Windows desktop programs based on their user interface. It is created on top of unmanaged Microsoft UI Automation API (Windows Automation API 3.0 - the latest version of UI Automation)..

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

<b>Events</b><br>
You can register a handler to be called when certain events occur. You can register a handler for edit element text changed or text selection changed event, combo box selected item changed event, check box checked event and so on.<br>
Suppose you already have a myButton object and you want to show a message box when the button is clicked:<br>
<i><b>myButton.ClickedEvent += (sender) => MessageBox.Show(sender.Text + " was clicked");</b></i>
<br><br>
<b>Browser automation</b><br>
You can automate Google Chrome and MS Edge browsers. For example navigate to a webpage and access a link:<br>
<i><b>ChromeBrowser browser = engine.NewChromeBrowser();<br>
browser.Navigate("wikipedia.org");<br>
UIDA_HyperLink link = browser.Hyperlink("English*", true);<br>
link.AccessLink();</b></i>

<br><b>Resources</b><br>
Samples: <a href="http://automationspy.freecluster.eu/Samples.Automation3.0.zip">http://automationspy.freecluster.eu/Samples.Automation3.0.zip</a>
<br><br>
CHM Help file: <a href="http://automationspy.freecluster.eu/UIDeskAutomation/UIDeskAutomation.Automation3.0_Help.chm">http://automationspy.freecluster.eu/UIDeskAutomation/UIDeskAutomation.Automation3.0_Help.chm</a>.<br>
When you open the help file uncheck the "Always ask before opening this file" checkbox so you can view the content.
<br><br>
UIDeskAutomationSpy tool: <a href="http://automationspy.freecluster.eu/UIDeskAutomation/UIDeskAutomationSpy.Automation3.0.zip">http://automationspy.freecluster.eu/UIDeskAutomation/UIDeskAutomationSpy.Automation3.0.zip</a>.<br>
If you encounter difficulties when opening this tool, right click on the spy executable file, choose Properties, check "Unblock" checkbox at the bottom of the dialog box, then Apply and OK.
