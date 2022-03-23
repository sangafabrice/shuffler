# **The shuffler component**

##### Author: Fabrice Sanga
<br/>
<br/>

The component is mainly built around the `RandomSelect()` function that selects an element from a list of objects randomly, as its name suggests. The selection proceeds similarly to a media player in shuffle mode. It avoids selecting an item that it had already picked til every item in the list is marked. Then the process restarts.

It works on any type of list, but for illustration purposes, it is restricted to Photos App Tiles, Desktop and Lockscreen Backgrounds, and the default Logon Picture.

An example of use:
```vbscript
With CreateObject("CustomUI.Shuffler")
    .WorkDir = "\Path\to\Example"           '(1)
    .Shuffle "DesktopBG"                    '(2)
    .Shuffle "LockScreenBG"                 
    .Shuffle "LogonPicture"                 
    .Shuffle "PhotosTile"                   
    .RefreshStartMenu                       '(3)
End With
```
**(1)** `WorkDir` is the root directory

**(2)** `Shuffle` is a sub that shuffles the items of the `DesktopBG` subfolder of the root

**(3)** `RefreshStartMenu` is a feature sub that makes the change available immediately and is needed for `LogonPicture` and `PhotosTile`
<br/>

The filesystem:
```
    Example
        |--DesktopBG
        |--LockScreenBG
        |--LogonPicture
        |--PhotosTile
```
`DesktopBG`, `LockScreenBG`, `PhotosTile` contain image objects.
`PhotosTile` contain subfolders of square images of `32px`, `40px`, `48px`, `192px` and `448px` size. They are different sized-images of the same picture. The figure shows an example of those images.
<br/>
<br/>
![](https://drive.google.com/uc?export=view&id=1qeUPHuRFGBPCtsyyAjGjoMOjcq_-Vchr)