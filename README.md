# ucAniGifEx v1.0
ucAniGifEx Animated Gif Control

![ss](https://github.com/user-attachments/assets/ad8fcf31-6807-4193-9550-4b00f8b19a91)

**ucAniGifExTest.twinproj** - Project with test form shown in the picture above.

**AniGifEx.twinproj** - ActiveX control project to build .ocx versions for VB6/VBA/etc

See [Releases](https://github.com/fafalone/ucAniGifEx/releases) for binaries.

```
/********************************************************************************
    ucAniGifEx v1.0
    by Jon Johnson, ported from Windows SDK WicAnimatedGif example, with
         corrections, additional features, and conversion to control.
    (c) 2025. Distributed under the MIT License.
    
    This is a followup to my ucAniGif control, which while it had very 
    simple code, there were a number of flaws. Some were mistakes that could
    be fixed, but others were limitations of the underlying GdiPlus code or
    impractical to fix in the shell interface (I already was resorting to
    patching assembly instructions in memory to fix the slow playback bug!).
    
    Many gifs did not play right, or at all. And very limited extras.
    
    While the code here is significantly more complex, this control supports 
    every gif I've found, but it's just as simple to use from the outside 
    while adding some basic enhancements.
        
    Features:
      -Multiple options to display a gif: 
         StartupFile property - Loads and plays a gif file from disk on load
         DisplayGifFromFile - Loads and plays a gif file from disk on call
         DisplayGifFromResource - Loads and plays a gif file from the project
              resources. Specify the ID and group; LoadResData is used.
         DisplayGifFromMemory - Play direct from a byte array containing a 
              complete gif file.
      
      -Paused property can stop/resume playback, StopPlayback command
      
      -Will be transparent to the control BackColor; it's recommended you set
       this before playing, i.e. ucAniGifEx1.BackColor = Me.BackColor, then play.
      
      -Gif is scaled to the control size by default.
      
      -PreserveAspectRatio option, with centering in the frame
      
      -SizeToFit - Resizes the control to the gif size. Scaling is disabled.
      
      -LoopControl - Override the gifs default to force infinite loops or only
          play once.
      
      -LoopEnd, PlaybackEnd, and FrameStep events. Note: FrameStep is disabled
          by default, to receive it, set EnablePerFrameEvents to True.
      
      -AdvanceByOneFrame - If paused, advance by a single frame without resuming.
      
      -FrameCount, FrameIndex, ImageWidth, ImageHeight properties.
    
    Requirements:
      Windows 7 or newer
      
    Known issues:
      twinBASIC currently displays a continuable exception error in the IDE. This
      does not seem to be impacting the control. A bug report has been filed.
      
    Change log: 
      (08 Apr 2025, v1.0) - Initial release.
      
    Thanks
      Special thanks to bahbahbah (whose code I wound up using), jpbro, and 
      The trick for helping me sort through some scaling issues.
      And wqweto for helping me out with some C++ issues to get the original
      SDK example working so I could compare.
*********************************************************************************/
```
