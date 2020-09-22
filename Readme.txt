Region Maker (Updated)
======================

This is a VERY useful tool for all of those VB programmers that want to reshape their forms. It takes an ordinary bitmap with a single background colour, and processes the image to make a region data file. The new file can easily be integrated in to your VB program via the module functions. The tool and code are very, very easy to use, although there are no comments in the code to tell you how to use it. All of that is mentioned here.

The updates are:
> Added header to make sure that the loaded file is a region data file
> Fixed small bug that keeps the preview window open if you close the Region Maker form
> Added an example region

The Tool
========

> Load Picture
	Click this to begin. Select the picture you wish to use.
> Selecting a background colour
	When you move the mouse around the image, the first colour box will change colour. When you see the colour that you want to appear completely transparent on the form, click the mouse. The selected colour will appear in the second box.
> Include option
	Normally, the colour you select is the colour that you do not want to see. When the Include checkbox is selected, everything but the selected colour disappears.
> Preview
	When you are satisfied with the colour selection, click preview. For large images, this will take a while, so just wait patiently. Eventually a new window will pop up in the shape of the bitmap. You can move it about by dragging any part of it that is visible. To get rid of the window, just switch focus back to the Region Maker.
> Save
	After everything is complete, you can save the region so that it can be read by the module functions at any time in the future.
> Load Region
	This is for testing purposes only, but it allows you to open up a previously saved region data file to preview.

The RegionManagement.bas Module
===============================

> ProjectRegion
	This places RegionData on to the form Window. You must select the "None" border style on the destination form.
> CreateRegion
	Creates region data from a bitmap or a form's picture. The parameters are self explanatory if you have read the instructions for the tool.
> SaveRegion and LoadRegion
	Saves or loads region data in to or from a file.



Important copyright notes:

WindowMover.cls may be used in any of your programs, whether they are for commercial use or not.
RegionManagement.bas may be used in any non-commercial programs freely, but if you wish to use it in commercial products, you must get my permission first.
The modules may only be modified if you are using them for personal use only.


Thankyou for reading this, and I hope my code is useful for you.

Craig Bonathan
craigthesnowman@yahoo.co.uk

CB3 Software
cb3software@fire-bug.co.uk
www.cb3software.fire-bug.co.uk