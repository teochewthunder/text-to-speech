# Text-To-Speech

This utlizes Microsoft's built-in VBScript engine, most notably the *sapi.spvoice* object and the *Speak* method.

First create the string you wish to be spoken. The example here has used Timer and Randomizer functions, but any string will do.

Then use the *sapi.spvoice* object and the *Speak* method. This is at the very end of the script.

To have this script execute whenever you start the computer, place the vbs file in *C:\Users\[your user name here]\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup*
