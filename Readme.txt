		+---------------------------+
		|   TRANS FUNCTION README   |
		+---------------------------+

		    +-- INTRODUCTION: --+

 "Trans" is a simple function that I wrote mainly for a
game I'm making, but I decided it was too cool to not share
it with others ;)

 What it does is to paint a picture to another one, but with
a nice translucency effect! Very cool if you want a small,
quick function to use translucency in your games --
something that only the top games you buy had! But beware:
if you use this function in large or many images at the same
time, it may slow down the game, because of the huge
calculations involved!

		    +-- INSTRUCTIONS: --+

 Simply add the .bas file to your project to use it.
 All you need to do in order to activate the effect is to
call this function.

			+-- USAGE: --+

 Trans(Target As Long, Source As Long, X As Long, Y As
Long, Width As Long, Height As Long)

		      +-- ARGUMENTS: --+

+- Target:
 This is the hDC where the secondary image will be
drawn.
(You'll find this as a property of forms and picture
boxes.
 Just know that it refers to the image used in these
controls.)

+- Source:
 This one is the hDC (same as above) of the image that
will be drawn to the first one. Use black for the areas
that you don't wanna paint.

+- X:
 The X coordinate of the place where the image will be
drawn (in the first image, of course).

+- Y:
 The Y coordinate of the place where the image will be
drawn (in the first image, of course).

+- Width:
 The width (in pixels) of the image contained in the hDC
given in the Target. If you're using a Form or a Picture
Box's hDC, you can just set that control's ScaleMode to
3 - Pixel and use here its ScaleWidth.

+- Height:
 The height (in pixels) of the image contained in the hDC
given in the Target. Same as above to ScaleWidth, but use
ScaleHeight instead here.

		       +--EXAMPLE: --+

 Check the sample project included for a working demo.

		+--------------------------+
		|  Made by Jotaf98         |
		|  (João F. S. Henriques)  |
		|                          |
		|  E-mail me at            |
		|  jotaf98@hotmail.com     |
		+--------------------------+