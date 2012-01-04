Virtual Input

This Windows-specific .Net library provides a simple way to intercept key presses and mouse activity in any application.

While it is possible to think of various malicious uses, such as keylogging, this is not the intended purpose - and there are other ways the bad guys can do that. VirtualInput was originally created for use in a screen reader for blind people - a screen reader will speak aloud which keys are pressed, and will allow special key combinations to be used for reviewing the screen. Other uses could include macro applications, automated UI tests, and various productivity tools.

At present the library focuses on intercepting user input - a future extension would be to also simulate input - such as simulating keypresses, mouse movement, and clicking the mouse.

You can add the library to your project using Nuget - see http://nuget.org for more information.
