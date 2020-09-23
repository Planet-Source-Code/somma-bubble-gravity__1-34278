
Bubble-Gravity: an Interactive Screensaver
------------------------------------------

Why the hell did you do this?

I was just getting some ideas about making a side-view deathmatch game, so I made this program first to test things like gravity, mass, and force. I just thought I'd turn it into a screensaver while I was at it.


What is this? What do I do with it?

Copy bubble-gravity.scr into your windows directory (it could be windows/, winnt/, or some other path depending on your operating system), and Bubble-Gravity will appear as an option in your screen saver settings. Alternatively, you can just right-click the file and chose 'install'. 

Remember that this is a screensaver - you don't have to do anything if you dont want to. There is an *Auto-screensaver* mode that will automatically do things with the bubbles to make the screensaver interesting (by default, this is on as soon as the program starts). However, it is an *interactive* screensaver, so there are a few things you can control. 

This is a program that simulates force and gravity on massive objects (massive just means 'has mass'). There are bubbles in the window which are attracted by gravity, and will bounce and move around accordingly (in parabolic curves and stuff). Gravity in any direction can be simulated by changing the X-Gravity (1 and 0 on the numpad) and Y-Gravity (2 and 3 on the numpad) values. Yes, you can make the bubbles 'fall' upwards and side-to-side if you want. In addition to gravity, there is Wind. Unlike gravity, the effect of wind will depend on the mass of the individual bubbles. Wind is controlled by 7 and 4 (up/down), and 8 and 9 (left/right) on the numpad.

The window starts off with one bubble. To add more, press the plus (+) key, or let the *Auto-Screensaver* do it for you. New bubbles will have a random intial velocity and mass. To take bubbles away, press the minus (-) key (note that the Auto-Screensaver will keep creating bubbles until the set limit is reached, so you should turn it off before popping bubbles).
Cycle through each bubble by pressing [Page Up] and [Page Down]. The mass of that bubble will be displayed at the top, and the selected bubble will be shown in red. 
To move the bubble around, press the arrow keys, and a *force* will be applied to the bubble in that direction. Because it is a force, the acceleration the bubble experiences will depend on its mass. To change the bubble's mass, press the [Insert] or [Delete] key. To change the force (thrust) you apply, press the [Home] or [End] key. The bubble wont move much if the thrust is too low for its mass.

If you forget the controls, just hold the F1 button down, and a list of controls will be shown.


Hey!...hang on a minute...

Before you start complaining, I will tell you now that this is *not* very accurate for a simulation. Although there is no friction when the bubbles bounce of the walls, you might notice that they will sometimes lose energy (bounce a bit lower) and then gain that energy back the next bounce. If you know a bit of physics, and you read the source code, you'll know why.


An attempt at some legal stuff.

'Som's Bubble-Gravity' (Program) is protected by National and International Copyright laws. The user (that would be *you*) may make copies of this program so long as the copyright owner (that would be *me*) recieves full recognition and entitlements for the program. The user may decompile, reverse engineer, and use source code from this program, so long as the owner (me again) recieves recognition and entitlements for any derivative works produced. 
In other words, please don't rip off my ideas without giving me at least some recognition for them. Thanks!