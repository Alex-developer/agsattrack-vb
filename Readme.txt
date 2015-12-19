AGSatTrack
----------

(V1.20)
-------

First release.

(V1.22)
-------

1. The current views can now be saved when you exit the program either automatically by selecting the options from the File->Properties dialog box or by selecting the save button on the toolbar.

2. The Globe view can be displayed as the desktop wallpaper. To select this option open a globe view and press the 'D' button on the toolbar.

3. You can now adjust the rate at wich the windows are updated. This is set in the File->Properties and copied to each window as they are created. It can also be adjusted from each window.

4. The time skip buttons now act on each window independently. Prior to this version they acted on all windows.

5. The select satellite dialog box now has a warning if the Keplarian elements are out of date. You can select the number of days before this warning appears from the File->Properties.

6. It is now possible to update the elements via the internet. On the satellite selection dialog box there is an update button. This is also available from the Elements menu. The elements are updated in groups. Three are provided in the Internet Updates folder. Later versions will allow these groups to be edited.

7. The installation has been rebuilt using the Windows 2000 installer to try and fix problems with the installation failing on Windows2000.

8. Now includes SimpleTrack a text based tracking tool (Installed in the main directory).

9. Lots of bug fixes and small speed improvements.

(V1.23)
-------

1. NEW: Added ability to edit the Internet Update groups.

2. NEW: Changed the colour of the moon footprint to make it more visible.

3. BUG: Removed all of the unused menu options.

4. NEW: Changed the toolbar icons to make them more obvious ?? (If anybody want to design some better ones then please let me know).

5. NEW: Speeded up the tracking control with some internal changes to the way its loaded.

6. NEW: Added speech option to announce expected AOS of selected satellite and after AOS the current azimuth and elevation, I couldn't get this to work on win95 on my laptop but it was fine on the development machine (Win ME)?. (This can be turned off!)

7. NEW: Any Toolbar changes are now saved when the program is closed down. To reset the toolbars use the reset option in the toolbars cutomisation.

8. BUG: The refresh button was not updating the satellite positions, only redrawing the current positions. This has now been fixed.

9. NEW: Got rid of the flashing scroll bars. The scroll bars on a tracking window would flash.

10.NEW: Removed the 'Keps' date from the status bar.

11.NEW: Changed the way the Lat/Lon is displayed in the satellite label, now shows North and South rather than +/-

12.BUG: If more than one 'DX Report' window was displayed then after the first refresh they would all display the same data.

13.NEW: Re-arranged the 'Window properties' tabs.

14.BUG: The Aos time reported on the status bar would wander. This has been fixed.

15.BUG: A small flashing icon appears on the predictions window whilst it is being updated - This has been removed.

16.NEW: Added toolbar option to display the satellite groundtrack as either dots or joined up lines. To get these options to display you will have to reset the toolbars. To do this delete the 'toolbars.atb' file in the main program directory.

17.NEW: Added toolbar option to specify the time interval between groundtrack dots. To get these options to display you will have to reset the toolbars. To do this delete the 'toolbars.atb' file in the main program directory.

18,NEW: Option added to the Edit menu to reset the toolbars back to the installation configuration. This option will ONLY be available once you have followed the instructions in points 16 and 17 above.

19.NEW: The satellites, sun and moon can be displayed as images on the mercator maps rather than dots, this is selected from the window properties dialog.

20.NEW: Most of the window settings are now saved when you exits the program or save the current view. Previously only a few settings were saved.

21.BUG: You could occasionally crash the program by clicking on a satellite, this has been fixed.

22.BUG: If you change the map image it is now saved when the program exits.

23.BUG: When displaying the satellite details from the satellite selection a runtime error ocurred.


(V1.24)
-------

1. BUG: Neither the program options nor the window options dialog would allow a negative Lat/Lon to be entered. This has been fixed.

2. BUG: The date formats were not using the windows control panel settings for their format. This has been fixed(?).

3. BUG: Observer locations were not being ploteed correctly when looking at the 180 degree map. This has been fixed.

4. BUG: The Observer lat/lon were displayed in the window properties dialog box to 10 decimal places! This has now been rounded to 2 decimal places.

5. NEW: Added mutual observation. This allos for a second observer to be setup. The ground track will show in yellow when the satellite is available from both locations.

(V1.25)
-------

1. NEW: Added SGP4/SDP4 models. This is selected by right clicking on a satellite and selecting the model option. SGP has been added throught the AGSGP.DLL windows DLL (Written in VC++). This has been done as the SGP models are very processor intensive.

2. BUG: The tablist display has numerous bugs. These have been fixed.

3. NEW: Various internal changes to the models.

4. BUG: The timezone was not saved with the satellite views.

5. NEW: Added ability to create new element groups. This can be used, for example, to add the Shuttle and ISS to a groups when the shuttle is in orbit. This option is on the 'Elements' menu. To activate the option select the 'Reset toolbars' option from the 'Edit' menu and restart the program.

6. NEW: Updated the satellite icons, these now look much better.

7. NEW: The age of the elements in any group is displayed in the satellite selection dialog. The youngest, oldest and average age are displayed. If the average is greated than the value specified in the program 'Properties' then a warning is displayed indicating that the elements should be updated.

8. NEW: The Tablist now supports the right mouse click menu. Right click on a satellite name to display the menu.

9. BUG: When changing to Daylight saving time the max elevation time on the horizon view showed the wrong time.

10.New: Added print preview option for map and prediction windows.

11.BUG: On Windows 2000 the internet update reports that you are never connected to the internet. Until I find the cause I have added an Ignore option which will attempt the download in any case.

12.NEW: The method for adding satellited to the OCX control has been altered to make it much easier, you now send it the raw keps data and it converts it internally.

(Outstanding Bugs)
------------------

1. The mercator maps do not display correctly as a wallpaper.

2. If you display the satellite data on more than one window then the data toggles between each form.

3. If you create more than one globe view and they update at the same instant the views can be corrupted.

4. The toolbars ONLY update whn switching windows if you click on the map in the window. Clicking on the titlebar will NOT update rhe toolbars. If you switch to a window and the toolbar is greyed out or does not have thr correct toolbar buttons enabled then click on the map and the toolbar will update.

5. Not really a bug but the program uses a massive amount of memory and I am not sure why! (About 20Mb on Win2k)

(Future Enhancements)
---------------------

1. Ability to have more that one satellite on the globe view.

2. Rewrite the globe view to support DirectX. This will allow for panning, rotation and zooming to be added. (I am currently learning DirectX).


