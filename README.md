Latin Macron Macros for Word 2003
======

These are some simple macros you can use to add buttons to a toolbar in Word (and probably other Office applications) so that you can get one-click access to adding Macron-enabled text to your documents.

[![No Maintenance Intended](http://unmaintained.tech/badge.svg)](http://unmaintained.tech/)

Installation Instructions - or on [My Site with Pictures!](https://yaleman.org/2012/03/07/latin-macrons-in-word-2003/)
-----

* Download the file - this is all the macro code.
* Open Word 2003.
* From the main screen, press the key combination *Alt-F11*.
* Ensure you're working on the *Normal* template, by double-clicking on *Normal* then *Microsoft Word Objects* then *ThisDocument*.
* Right click on the *Normal* title and click *Import File...*
* Select the file and click *Open*.
* This should import the file into a *Module* called *Module1* or similar, double click on *Modules* then double click on the newly created module item.
* To rename it so we know what we're dealing with, go down to the module properties (which should be in the bottom left) and click next to *(Name)*.
* Delete what's in the box and rename it to *Macrons*.
* Hit enter and it'll update the name.
* Click on the *Save* icon in the top of the window to confirm the changes to the Normal template.
* Close the *Microsoft Visual Basic* window and you'll be back in *Word*
* Click on *View* -> *Toolbars* -> *Customize...*
* Go to the *Toolbars* tab and create a new toolbar by clicking on *New...* and giving it a name. Ensure you've set *Make toolbar available to:* to *Normal.dot*.
* Change to the *Commands* tab.
* Scroll down in the left hand window and find *Macros* then click on it.

For each of the *Commands* in the right hand side which start with *Normal.Macrons.*:  

* Drag the option up onto your new toolbar. If you put it in the wrong spot, just click on it and drag it to another spot.
* Right click on the option and give it a more useful name. I tend to go with just the letter (or accent-free representation of it) to take up less room.
* Set the style to *Text Only (Always)* by clicking on that option.

At this point you're done, and you can click on a toolbar button to insert the character in your document!

If you want to set a keyboard shortcut for your macros:

* Click on *Keyboard...* from the *Customize* screen.  
* Select *Macros* on the left, then find the applicable macro on the right and click on it.
* Click in the box under *Press new shortcut key:*
* Press the key combination you want to use.
* If you are replacing an existing key combination, it'll show what you're replacing next to *Currently assigned to:*, otherwise it'll say *[unassigned]*
* Click on Assign and it's your new shortcut key.

If you want to later remove the shortcut, come back to the screen, find the macro and click *Remove*  
