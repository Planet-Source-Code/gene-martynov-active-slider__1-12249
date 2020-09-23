Why should you use that standard slider? It is not nice any more.
Try to use this one instead.
This is a horizontal slider ActiveX. You can drag the pointer from minimum to maximum positions
with your mouse. Or you can click on any visible part of slider to set pointer at this
position. Or you can set position programatically. Includes PositionChanged Event. Control can
be transparent or opaque. You can configure your slider in the way you like it most. 
Source codes and demo project included.


The properties you can use:
Min (Long) - sets/returns minimum value for slider.

Max (Long) - sets/returns maximum value for slider.

CurPosition (Long) - sets/returns current slider position.

ImageSlider - sets a picture to use with slider. If you use ImageLeft then make sure that
width and height of both images are the same.

ImagePointer - sets a picture to use with pointer.

ImageLeft - sets the picture to use on the left side of pointer. If this picture is not set
no changes made on the left side of the pointer. What I mean - slider will look the same on
both sides of pointer. Make sure the height and width of this image is the same as Slider Image.
If it is not you will get some garbage on the result picture.

BackStyle (Integer)- sets/returns whether slider is transparent or opaque. Transparancy is
specified by MaskColor. Can be 0 - transparent, or 1 - opaque.

MaskColor (OLE_Color/Long) - specifies the color on all images that will be used as
transparent. Make sure that you don't use this color on the parts that should not be
transparent. Also make sure that you use one color on all images for transparancy.

Style (Integer) - sets/returns whether slider uses discrete positions or not. Can be 0 -
analog, or 1 - discrete.

********* only in new version**********
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Added properties in new version:
DisabledStyle (Long) - sets how control is displayed when disabled. Can be 0 - show 
disabled control lighter, 1 - show disabled control darker.

DisabledIntense (Long) - sets the value that is added or subtracted to/from each pixel on
normal image, when control is disabled. Allowable value from 0 to 85. You can play with it
and adjust view of disabled control at desirable value.

Orientation (Long) - specifies the orientation of minimum and maximum on slider:
0 - minimum on the left side of slider, maximum - on right
1 - minimum on the right side of slider, maximum - on left
2 - minimum on the top side of slider, maximum - on bottom
3 - minimum on bottom side of slider, maximum - on top

All other properties are standard.


Events you can use:
PositionChanged(oldPosition as Long, newPosition as Long) - this event is fired every time
when the position of slider is changed. It means you don't need a timer to keep track of
position.

All other events are standard.


ENJOY!!!!!!!