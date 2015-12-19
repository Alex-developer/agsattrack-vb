HelpScribble project file.
10
Nyrk Terraynaq-0N8954
0
1
AGSatTrack


(c) 1999-2000 Alex Greenland
FALSE

C:\MYDOCU~1\MYPICT~1,C:\MYDOCU~1\DEVELOP\VISUAL~1\RADIO\SATELL~1
1
BrowseButtons()
0
FALSE
C:\My Documents\Develop\Visual Basic\Radio\Satellite Tracking\
FALSE
19
12
Scribble12
Acknowledgments




Writing



FALSE
29
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\froman\fprq2 Times New Roman;}}
{\colortbl\red0\green0\blue0;\red0\green0\blue255;}
\deflang2057\pard\plain\f2\fs32\cf1\b Acknowledgments\plain\f2\fs20\cf0 
\par 
\par \plain\f3\fs24 I am indebted to the following people for their help and support during the development of this software
\par 
\par Nic \endash  The long suffering 'Radio' widow.
\par \plain\f2\fs20\cf0 
\par Andrew Davidson - For all of the help testing.
\par 
\par Karl (G6ODT) - For all of the help testing, and producing the first program to use the OCX control.
\par 
\par Part of the code was ported from a Linux earthviewer (xEarth). The author of xEarth requires that the following notice appear in all software.
\par 
\par Copyright (C) 1989, 1990, 1993, 1994, 1995 Kirk Lauritz Johnson
\par 
\par Parts of the source code (as marked) are:
\par   Copyright (C) 1989, 1990, 1991 by Jim Frost
\par   Copyright (C) 1992 by Jamie Zawinski <jwz@lucid.com>
\par 
\par Permission to use, copy, modify and freely distribute xearth for
\par non-commercial and not-for-profit purposes is hereby granted
\par without fee, provided that both the above copyright notice and this
\par permission notice appear in all copies and in supporting
\par documentation.
\par 
\par 
\par And all of the others that have tested the software for me. Thankyou
\par }
15
Scribble15
Element Warning




Writing



FALSE
20
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}}
{\colortbl\red0\green0\blue0;\red0\green0\blue255;}
\deflang2057\pard\plain\f3\fs32\cf1\b Element Warning\plain\f3\fs20\cf0 
\par 
\par Element Set Age Warning
\par 
\par When a new element set file is loaded, the average element set age is computed.  The age is the difference between the current time and the epoch of the elset.  If the average age of the elsets is greater than the number of days specified in the program options, 14 by default, days, this dialog is displayed. It shows the average age and age of the oldest and youngest elsets, all in days.  
\par 
\par 
\par What does it mean?
\par 
\par Element sets are perishable items, just like fruits and vegetables.  If they are not used within a relatively short time, they are no longer valid for predictions.  The useful life of an element set depends on the period of the orbit.  Generally speaking, the higher the orbit the longer the elset can be used.  This is because the primary force perturbing an orbit is drag imparted by the earth's atmosphere.  Lower orbits (smaller periods) suffer more from drag than do higher orbits (larger periods).
\par 
\par For an elset with a period of less than 90 minutes, a three-week old elset is probably no longer good.  The same age for an elset with a period of 100 minutes is probably still good enough for predictions.
\par 
\par 
\par 
\par 
\par 
\par }
20
Scribble20
Overview




Writing



FALSE
31
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}}
{\colortbl\red0\green0\blue0;\red128\green0\blue0;\red0\green0\blue255;}
\deflang2057\pard\plain\f3\fs32\cf2\b Overview
\par \plain\f3\fs20\cf0 
\par \plain\f3\fs20\cf1 \{bmc SPLASH.bmp\}\plain\f3\fs20\cf0 
\par 
\par AGSatTrack allows earth orbiting satellites to be tracked. The program allows for several windows to be open, each displaying different satellites and different views.
\par 
\par \plain\f3\fs24\cf2 Main Features\plain\f3\fs20\cf0 
\par 
\par \pard\li200\fi-200{\*\pn\pnlvlblt\pnf1\pnindent200{\pntxtb\'b7}}\plain\f3\fs20\cf0 {\pntext\f1\'b7\tab}Multiple windows, each displaying a different view and different satellites.
\par {\pntext\f1\'b7\tab}0 and 180 degree centered earth mercator projections.
\par {\pntext\f1\'b7\tab}Horizon view, 360 degree panoramic view.
\par {\pntext\f1\'b7\tab}Orthographic view.
\par {\pntext\f1\'b7\tab}Tabular view.
\par {\pntext\f1\'b7\tab}Sun and moon plots.
\par {\pntext\f1\'b7\tab}DX available from satellite report, for radio amateurs.
\par {\pntext\f1\'b7\tab}Satellite footprint.
\par {\pntext\f1\'b7\tab}Satellite groundtrack, for next orbit.
\par {\pntext\f1\'b7\tab}Pass prediction report
\par {\pntext\f1\'b7\tab}Pass analysis report
\par {\pntext\f1\'b7\tab}Doppler shift corrections
\par {\pntext\f1\'b7\tab}Path loss
\par {\pntext\f1\'b7\tab}Selectable maps (User can alter the world and horizon maps)
\par {\pntext\f1\'b7\tab}Each window can have a different observer.
\par {\pntext\f1\'b7\tab}Keplarian element Editing
\par {\pntext\f1\'b7\tab}Internet based Keplarian element updating.
\par \pard\plain\f3\fs20\cf0 
\par 
\par 
\par }
30
Scribble30
Getting Started




Writing



FALSE
7
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}}
{\colortbl\red0\green0\blue0;\red128\green0\blue0;\red0\green128\blue0;\red0\green0\blue255;}
\deflang2057\pard\plain\f3\fs32\cf3\b Getting Started
\par \plain\f3\fs20\cf0 
\par When the program is first run after installation a new tracking view will be opened and a list of available satellites displayed. From this list select the satellites that you wish to track on the tracking view. Please refer to the \plain\f3\fs20\cf2\strike satellite selection\plain\f3\fs20\cf1 \{linkID=35\}\plain\f3\fs20\cf0  section of this help file for more details on using this facility.
\par 
\par }
35
Scribble35
Satellite Selection




Writing



FALSE
18
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}{\f4\fswiss\fprq2 System;}}
{\colortbl\red0\green0\blue0;\red0\green0\blue255;\red128\green0\blue0;\red0\green128\blue0;}
\deflang2057\pard\plain\f3\fs32\cf1\b Satellite Selection\plain\f3\fs20\cf0 
\par 
\par The satellite selection list is displayed when a new window is created or the \{button] is selected from the toolbar.
\par 
\par \plain\f3\fs20\cf2 \{bmc satsel.bmp\}\plain\f3\fs20\cf0 
\par 
\par The selection list will display all of the satellites available from the group selected, the group can be changed by using the drop down list
\par 
\par Upto 20 satellites can be selected from this list to be displayed on a tracking view. To select a satellite click on the tick box next to the satellite identifier (Norad id).
\par 
\par The list can be sorted by clicking on the column headers, by default the list is sorted by Norad Id.
\par 
\par The list displays the Norad Identifier, the satellite name and the date of the \plain\f3\fs20\cf3\strike Keplarian elements\plain\f3\fs20\cf2 \{linkID=35\}\plain\f3\fs20\cf0 . Ideally these elements should be updated weekly to ensure accurate predictions. The Update button will perform an update from the internet given a suitable 
\par 
\par The Update button will display the \plain\f3\fs20\cf3\strike Internet Update\plain\f3\fs20\cf2 \{linkID=37\}\plain\f3\fs20\cf0  dialog and allow the keplarian elements to be updated via the internet. For this to work you must have an internet connection already conneted. This feature will not currently work through a proxy server.
\par }
37
Scribble37
Internet Update




Writing



FALSE
13
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}{\f4\fswiss\fprq2 System;}}
{\colortbl\red0\green0\blue0;\red0\green0\blue255;\red0\green128\blue0;\red128\green0\blue0;}
\deflang2057\pard\plain\f3\fs32\cf1\b Internet Update\plain\f3\fs20\cf0 
\par 
\par This option will update the \plain\f3\fs20\cf2\strike Keplarian elements\plain\f3\fs20\cf3 \{linkID=60\}\plain\f3\fs20\cf0  via the Internet. For this to work you MUST be connected to the internet, no attempt will be made to establish a connection, and you MUST not have a Proxy server. Future versions will provide support for Proxy servers.
\par 
\par \plain\f3\fs20\cf3 \{bmc inetupdate.bmp\}\plain\f3\fs20\cf0 
\par 
\par It is possible to update the \plain\f3\fs20\cf2\strike Keplarian elements\plain\f3\fs20\cf3 \{linkID=60\}\plain\f3\fs20\cf0  in groups, three are provided with the installation. The drop down list provides access to the groups. Once a group has been selected pressing the Start button will update the element. If you are updating the elements from the Satellite selection then when the update is complete the list will be refreshed displaying the new elements. The elements that are retrieved are \plain\f3\fs20\cf2\strike two line format\plain\f3\fs20\cf3 \{linkID=39\}\plain\f3\fs20\cf0 .
\par 
\par It is recommended that the elements are updated at least once a week for accurate tracking.
\par 
\par }
39
Scribble39
2 Line Element Sets




Writing



FALSE
47
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss Courier New;}{\f4\froman Times New Roman;}}
{\colortbl\red0\green0\blue0;\red0\green0\blue255;}
\deflang2057\pard\plain\f2\fs32\cf1\b 2 Line Element Sets\plain\f2\fs20\cf0 
\par 
\par \plain\f4\fs24 Data for each satellite consists of three lines in the following format:
\par \pard\tx0\tx959\tx1918\tx2877\tx3836\tx4795\tx5754\tx6713\tx7672\tx8631\plain\f3\fs20 AAAAAAAAAAAAAAAAAAAAAAAA
\par 1 NNNNNU NNNNNAAA NNNNN.NNNNNNNN +.NNNNNNNN +NNNNN-N +NNNNN-N N NNNNN
\par 2 NNNNN NNN.NNNN NNN.NNNN NNNNNNN NNN.NNNN NNN.NNNN NN.NNNNNNNNNNNNNN
\par \pard\plain\f4\fs24 Line 0 is a twenty-four character name (to be consistent with the name length in the NORAD SATCAT).
\par Lines 1 and 2 are the standard Two-Line Orbital Element Set Format identical to that used by NORAD and NASA. The format description is:
\par \plain\f4\fs24\b 
\par Line 1\tab \plain\f4\fs24 \tab 
\par \plain\f4\fs24\b Column\tab Description\tab 
\par \plain\f4\fs24 01\tab Line Number of Element Data\tab 
\par 03-07\tab Satellite Number\tab 
\par 08\tab Classification (U=Unclassified)\tab 
\par 10-11\tab International Designator (Last two digits of launch year)\tab 
\par 12-14\tab International Designator (Launch number of the year)\tab 
\par 15-17\tab International Designator (Piece of the launch)\tab 
\par 19-20\tab Epoch Year (Last two digits of year)\tab 
\par 21-32\tab Epoch (Day of the year and fractional portion of the day)\tab 
\par 34-43\tab First Time Derivative of the Mean Motion\tab 
\par 45-52\tab Second Time Derivative of Mean Motion (decimal point assumed)\tab 
\par 54-61\tab BSTAR drag term (decimal point assumed)\tab 
\par 63\tab Ephemeris type\tab 
\par 65-68\tab Element number\tab 
\par 69\tab Checksum (Modulo 10) (Letters, blanks, periods, plus signs = 0; minus signs = 1)\tab 
\par \pard\qc\plain\f4\fs24 
\par \pard\plain\f4\fs24\b Line 2\tab \plain\f4\fs24 \tab 
\par \plain\f4\fs24\b Column\tab Description\tab 
\par \plain\f4\fs24 01\tab Line Number of Element Data\tab 
\par 03-07\tab Satellite Number\tab 
\par 09-16\tab Inclination [Degrees]\tab 
\par 18-25\tab Right Ascension of the Ascending Node [Degrees]\tab 
\par 27-33\tab Eccentricity (decimal point assumed)\tab 
\par 35-42\tab Argument of Perigee [Degrees]\tab 
\par 44-51\tab Mean Anomaly [Degrees]\tab 
\par 53-63\tab Mean Motion [Revs per day]\tab 
\par 64-68\tab Revolution number at epoch [Revs]\tab 
\par 69\tab Checksum (Modulo 10)\tab 
\par \pard\qc\plain\f4\fs24 All other columns are blank or fixed.
\par \pard\plain\f4\fs24 Example:
\par \pard\tx0\tx959\tx1918\tx2877\tx3836\tx4795\tx5754\tx6713\tx7672\tx8631\plain\f3\fs20 NOAA 14                 
\par 1 23455U 94089A   97320.90946019  .00000140  00000-0  10191-3 0  2621
\par 2 23455  99.0090 272.6745 0008546 223.1686 136.8816 14.11711747148495
\par \pard\plain\f2\fs20\cf0 
\par }
40
Scribble40
Setting up the program




Writing



FALSE
9
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}}
{\colortbl\red0\green0\blue0;\red0\green0\blue255;\red0\green128\blue0;\red128\green0\blue0;}
\deflang2057\pard\plain\f3\fs32\cf1\b Setting up the program\plain\f3\fs20\cf0 
\par 
\par For predictions to be accurate you must setup your location, timezone and ensure that the Keplarian elements are up to date.
\par 
\par From the File menu select the \plain\f3\fs20\cf2\strike Properties\plain\f3\fs20\cf3 \{linkID=50\}\plain\f3\fs20\cf0  menu item. This will display the main program options. From here you can set your location and timezone. These options are duplicated in each new view that is created, and can be changed from the view options. Thus it is possible to have two different views with different observers.
\par 
\par }
50
Scribble50
Main Options




Writing



FALSE
8
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}}
{\colortbl\red0\green0\blue0;\red0\green0\blue255;}
\deflang2057\pard\plain\f3\fs32\cf1\b Main Options\plain\f3\fs20\cf0 
\par 
\par The main options dialog contains X tabs
\par 
\par 
\par }
60
Scribble60
Keplarian Elements
Keps;Keplarian



Writing



FALSE
125
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}}
{\colortbl\red0\green0\blue0;\red0\green0\blue255;}
\deflang2057\pard\plain\f3\fs32\cf1\b Keplarian Elements
\par \plain\f3\fs20\cf0 
\par \plain\f3\fs20 
\par Satellite Orbital Elements are numbers that tell us the orbit of each satellite. 
\par Elements for common satellites are distributed through amateur radio bulletin 
\par boards, and other means. 
\par 
\par 
\par \plain\f3\fs28\cf1 The Seven (or Eight) Keplerian Elements\plain\f3\fs20 
\par 
\par Seven numbers are required to define a satellite orbit. This set of seven numbers is called the satellite orbital elements, or sometimes "Keplerian" elements (after Johann Kepler [1571-1630]), or just elements. These numbers define an ellipse, orient it about the earth, and place the satellite on the ellipse at a particular time. In the Keplerian model, satellites orbit in an 
\par ellipse of constant shape and orientation. 
\par The real world is slightly more complex than the Keplerian model, and tracking programs compensate for this by introducing minor corrections to the Keplerian model. These corrections are known as perturbations. The perturbations that amateur tracking programs know about are due to the lumpiness of the earth's gravitational field (which luckily you don't have to specify), and the "drag" on the satellite due to atmosphere. Drag becomes an optional eighth orbital 
\par element. 
\par Orbital elements remain a mystery to most people. This is due I think first to the aversion many people (including me) have to thinking in three dimensions, and second to the horrible names the ancient astronomers gave these seven simple numbers and a few related concepts. To make matters worse, sometimes several different names are used to specify the same number. Vocabulary is the hardest part of celestial mechanics!
\par  
\par The basic orbital elements are...
\par   Epoch 
\par   Orbital Inclination 
\par   Right Ascension of Ascending Node (R.A.A.N.) 
\par   Argument of Perigee 
\par   Eccentricity 
\par   Mean Motion 
\par   Mean Anomaly 
\par   Drag (optional) 
\par 
\par The following definitions are intended to be easy to understand. More rigorous definitions can be found in almost any book on the subject. I've used aka as an abbreviation for "also known as" in the following text. 
\par 
\par \plain\f3\fs28\cf1 Epoch\plain\f3\fs20 
\par 
\par [aka "Epoch Time" or "T0"] 
\par 
\par A set of orbital elements is a snapshot, at a particular time, of the orbit of a satellite. Epoch is simply a number which specifies the time at which the snapshot was taken.
\par  
\par \plain\f3\fs28\cf1 Orbital Inclination\plain\f3\fs20 
\par 
\par [aka "Inclination" or "I0"] 
\par 
\par The orbit ellipse lies in a plane known as the orbital plane. The orbital plane always goes through the center of the earth, but may be tilted any angle relative to the equator. Inclination is the angle between the orbital plane and the equatorial plane. By convention, inclination is a number between 0 and 180 degrees. 
\par Some vocabulary: Orbits with inclination near 0 degrees are called equatorial orbits (because the satellite stays nearly over the equator). Orbits with inclination near 90 degrees are called polar (because the satellite crosses over the north and south poles). The intersection of the equatorial plane and the orbital plane is a line which is called the line of nodes. More about that 
\par later.
\par  
\par \plain\f3\fs24\cf1 Right Ascension of Ascending Node\plain\f3\fs20 
\par 
\par [aka "RAAN" or "RA of Node" or "O0", and occasionally called "Longitude of 
\par Ascending Node"] 
\par 
\par RAAN wins the prize for most horribly named orbital element. Two numbers orient the orbital plane in space. The first number was Inclination. This is the second. After we've specified inclination, there are still an infinite number of orbital planes possible. The line of nodes can poke out the anywhere along the equator. If we specify where along the equator the line of nodes pokes out, we will have the orbital plane fully specified. The line of nodes pokes out two places, of course. We only need to specify one of them. One is called the ascending node (where the satellite crosses the equator going from south to north). The other is called the descending node (where the satellite crosses the equator going from north to south). By convention, we specify the location of the ascending node. Now, the earth is spinning. This means that we can't use the common latitude/longitude coordinate system to specify where the line of nodes points. Instead, we use an astronomical coordinate system, known as the right ascension / declination coordinate system, which does not spin with the earth. Right ascension is another fancy word for an angle, in this case, an angle measured in the equatorial plane from a reference point in the sky where right ascension is defined to be zero. Astronomers call this point the vernal equinox. Finally, "right ascension of ascending node" is an angle, measured at the center 
\par of the earth, from the vernal equinox to the ascending node. I know this is getting complicated. Here's an example. Draw a line from the center of the earth to the point where our satellite crosses the equator (going from south to north). If this line points directly at the vernal equinox, then RAAN = 0 degrees. 
\par By convention, RAAN is a number in the range 0 to 360 degrees. I used the term "vernal equinox" above without really defining it. If you can tolerate a minor digression, I'll do that now. Teachers have told children for years that the vernal equinox is "the place in the sky where the sun rises on the first day of Spring". This is a horrible definition. Most teachers, and students, have no idea what the first day of spring is (except a date on a calendar), and no idea why the sun should be in the same place in the sky on that date every year. 
\par You now have enough astronomy vocabulary to get a better definition. Consider the orbit of the sun around the earth. I know in school they told you the earth orbits around the sun, but the math is equally valid either way, and it suits our needs at this instant to think of the sun orbiting the earth. The orbit of the sun has an inclination of about 23.5 degrees. (Astronomers don't usually call this 23.5 degree angle an 'inclination', by the way. They use an infinitely more obscure name: The Obliquity of The Ecliptic.) The orbit of the sun is divided (by humans) into four equally sized portions called seasons. The one called Spring begins when the sun pops up past the equator. In other words, the first day of Spring is the day that the sun crosses through the equatorial plane going from South to North. We have a name for that! It's the ascending node of the Sun's orbit. So finally, the vernal equinox is nothing more than the ascending node of the Sun's orbit. The Sun's orbit has RAAN = 0 simply because we've defined the Sun's ascending node as the place from which all ascending nodes are measured. The RAAN of your satellite's orbit is just the angle (measured at the center of the earth) between the place the Sun's orbit pops up past the equator, and the place your satellite's orbit pops up past the equator. 
\par 
\par \plain\f3\fs24\cf1 Argument of Perigee\plain\f3\fs20 
\par 
\par [aka "ARGP" or "W0"]
\par  
\par Argument is yet another fancy word for angle. Now that we've oriented the orbital plane in space, we need to orient the orbit ellipse in the orbital plane. We do this by specifying a single angle known as argument of perigee. A few words about elliptical orbits... The point where the satellite is closest to the earth is called perigee, although it's sometimes called periapsis or perifocus. We'll call it perigee. The point where the satellite is farthest from earth is called apogee (aka apoapsis, or apifocus). If we draw a line from perigee to apogee, this line is called the line-of-apsides. (Apsides is, of course, the plural of apsis.) I know, this is getting complicated again. 
\par Sometimes the line-of-apsides is called the major-axis of the ellipse. It's just a line drawn through the ellipse the "long way". The line-of-apsides passes through the center of the earth. We've already identified another line passing through the center of the earth: the line of 
\par nodes. The angle between these two lines is called the argument of perigee. Where any two lines intersect, they form two complimentary angles, so to be specific, we say that argument of perigee is the angle (measured at the center of the earth) from the ascending node to perigee. 
\par Example: When ARGP = 0, the perigee occurs at the same place as the ascending node. That means that the satellite would be closest to earth just as it rises up over the equator. When ARGP = 180 degrees, apogee would occur at the same place as the ascending node. That means that the satellite would be farthest from earth just as it rises up over the equator. 
\par By convention, ARGP is an angle between 0 and 360 degrees. 
\par 
\par \plain\f3\fs24\cf1 Eccentricity\plain\f3\fs20 
\par 
\par [aka "ecce" or "E0" or "e"] 
\par 
\par This one is simple. In the Keplerian orbit model, the satellite orbit is an ellipse. Eccentricity tells us the "shape" of the ellipse. When e=0, the ellipse is a circle. When e is very near 1, the ellipse is very long and skinny. (To be precise, the Keplerian orbit is a conic section, which can be either an ellipse, which includes circles, a parabola, a hyperbola, or a straight line! But here, we are only interested in elliptical orbits. The other kinds of orbits are not used for satellites, at least not on purpose, and tracking programs typically aren't programmed to handle them.) For our purposes, eccentricity must be in the range 0 <= e < 1. 
\par 
\par \plain\f3\fs24\cf1 Mean Motion\plain\f3\fs20 
\par 
\par [aka "N0"] (related to "orbit period" and "semimajor-axis") 
\par 
\par So far we've nailed down the orientation of the orbital plane, the orientation of the orbit ellipse in the orbital plane, and the shape of the orbit ellipse. Now we need to know the "size" of the orbit ellipse. In other words, how far away is the satellite? Kepler's third law of orbital motion gives us a precise relationship between the speed of the satellite and its distance from the earth. Satellites that are close to the earth orbit very quickly. Satellites far away orbit slowly. This means that we could accomplish the same thing by specifying either the speed at which the satellite is moving, or its distance from the earth! Satellites in circular orbits travel at a constant speed. Simple. We just specify that speed, and we're done. Satellites in non-circular (i.e., eccentricity > 0) orbits move faster when they are closer to the earth, and slower when they are farther away. The common practice is to average the speed. You could call this number average speed", but astronomers call it the "Mean Motion". Mean Motion is usually given in units of evolutions per day. In this context, a revolution or period is defined as the time from one perigee 
\par to the next. Sometimes "orbit period" is specified as an orbital element instead of Mean Motion. Period is simply the reciprocal of Mean Motion. A satellite with a Mean Motion of 2 revs per day, for example, has a period of 12 hours. Sometimes semi-major-axis (SMA) is specified instead of Mean Motion. SMA is one-half the length (measured the long way) of the orbit ellipse, and is 
\par directly related to mean motion by a simple equation. Typically, satellites have Mean Motions in the range of 1 rev/day to about 16 rev/day. 
\par 
\par \plain\f3\fs24\cf1 Mean Anomaly\plain\f3\fs20 
\par 
\par [aka "M0" or "MA" or "Phase"] 
\par 
\par Now that we have the size, shape, and orientation of the orbit firmly established, the only thing left to do is specify where exactly the satellite is on this orbit ellipse at some particular time. Our very first orbital element (Epoch) specified a particular time, so all we need to do now is specify where, on the ellipse, our satellite was exactly at the Epoch time. Anomaly is yet another astronomer-word for angle. Mean anomaly is simply an angle that marches uniformly in time from 0 to 360 degrees during one revolution. It is defined to be 0 degrees at perigee, and therefore is 180 degrees at apogee. If you had a satellite in a circular orbit (therefore moving at constant speed) and you stood in the center of the earth and measured this angle from perigee, 
\par you would point directly at the satellite. Satellites in non-circular orbits move at a non-constant speed, so this simple relation doesn't hold. This relation does hold for two important points on the orbit, however, no matter what the eccentricity. Perigee always occurs at MA = 0, and apogee always occurs at MA = 180 degrees. It has become common practice with radio amateur satellites to use Mean Anomaly to schedule satellite operations. Satellites commonly change modes or turn on or off at specific places in their orbits, specified by Mean Anomaly. 
\par Unfortunately, when used this way, it is common to specify MA in units of 256ths of a circle instead of degrees! Some tracking programs use the term "phase" when they display MA in these units. It is still specified in degrees, between 0 and 360, when entered as an orbital element. 
\par Example: Suppose Oscar-99 has a period of 12 hours, and is turned off from Phase 240 to 16. That means it's off for 32 ticks of phase. There are 256 of these ticks in the entire 12 hour orbit, so it's off for (32/256)x12hrs = 1.5 hours. Note that the off time is centered on perigee. Satellites in highly eccentric orbits are often turned off near perigee when they're moving the fastest, and 
\par therefore difficult to use. 
\par 
\par \plain\f3\fs24\cf1 Drag\plain\f3\fs20 
\par 
\par [aka "N1"] 
\par 
\par Drag caused by the earth's atmosphere causes satellites to spiral downward. As they spiral downward, they speed up. The Drag orbital element simply tells us the rate at which Mean Motion is changing due to drag or other related effects. Precisely, Drag is one half the first time derivative of Mean Motion. Its units are revolutions per day per day. It is typically a very small number. Common values for low-earth-orbiting satellites are on the order of 10^-4. Common values for high-orbiting satellites are on the order of 10^-7 or smaller. 
\par Occasionally, published orbital elements for a high-orbiting satellite will show a negative Drag! At first, this may seem absurd. Drag due to friction with the earth's atmosphere can only make a satellite spiral downward, never upward. There are several potential reasons for negative drag. First, the measurement which produced the orbital elements may have been in error. It is common to estimate orbital elements from a small number of observations made over a short 
\par period of time. With such measurements, it is extremely difficult to estimate Drag. Very ordinary small errors in measurement can produce a small negative drag. The second potential cause for a negative drag in published elements is a little more complex. A satellite is subject to many forces besides the two we have discussed so far (earth's gravity, and atmospheric drag). Some of these forces (for example gravity of the sun and moon) may act together to cause a satellite 
\par to be pulled upward by a very slight amount. This can happen if the Sun and Moon are aligned with the satellite's orbit in a particular way. If the orbit is measured when this is happening, a small negative Drag term may actually provide the best possible 'fit' to the actual satellite motion over a *short* period of time. 
\par You typically want a set of orbital elements to estimate the position of a satellite reasonably well for as long as possible, often several months. Negative Drag never accurately reflects what's happening over a long period of time. Some programs will accept negative values for Drag, but I don't approve of them. Feel free to substitute zero in place of any published negative Drag 
\par value. 
\par 
\par \plain\f3\fs24\cf1 Other Satellite Parameters\plain\f3\fs20 
\par 
\par All the satellite parameters described below are optional. They allow tracking programs to provide more information that may be useful or fun. 
\par 
\par \plain\f3\fs24\cf1 Epoch Rev\plain\f3\fs20 
\par 
\par [aka "Revolution Number at Epoch"] 
\par 
\par This tells the tracking program how many times the satellite has orbited from the time it was launched until the time specified by "Epoch". Epoch Rev is used to calculate the revolution number displayed by the tracking program. Don't be surprised if you find that orbital element sets which come from NASA have incorrect values for Epoch Rev. The folks who compute satellite orbits don't tend to pay a great deal of attention to this number! At the time of this 
\par writing [1989], elements from NASA have an incorrect Epoch Rev for Oscar-10 and Oscar-13. Unless you use the revolution number for your own bookeeping purposes, you needn't worry about the accuracy of Epoch Rev. 
\par 
\par \plain\f3\fs24\cf1 Attitude\plain\f3\fs20 
\par 
\par [aka "Bahn Coordinates"] 
\par 
\par The spacecraft attitude is a measure of how the satellite is oriented in space. Hopefully, it is oriented so that its antennas point toward you! There are several orientation schemes used in satellites. The Bahn coordinates apply only to spacecraft which are spin-stablized. Spin-stabilized satellites maintain a constant inertial orientation, i.e., its antennas point a fixed direction in 
\par space (examples: Oscar-10, Oscar-13). 
\par The Bahn coordinates consist of two angles, often called Bahn Latitude and Bahn Longitude. These are published from time to time for the elliptical-orbit amateur radio satellites in various amateur satellite publications. Ideally, these numbers remain constant except when the spacecraft controllers are re-orienting the spacecraft. In practice, they drift slowly. For highly elliptical orbits (Oscar-10, Oscar-13, etc.) these numbers are usually in the vicinity of: 0,180. This means that the antennas point directly toward earth when the satellite is at apogee. 
\par These two numbers describe a direction in a spherical coordinate system, just as geographic latitude and longitude describe a direction from the center of the earth. In this case, however, the primary axis is along the vector from the satellite to the center of the earth when the satellite is at perigee. An excellent description of Bahn coordinates can be found in Phil Karn's "Bahn 
\par Coordinates Guide". 
\par 
\par 
\par 
\par 
\par 
\par }
5000
Scribble5000
Select Satellite OK




Writing



FALSE
6
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}}
{\colortbl\red0\green0\blue0;}
\deflang2057\pard\plain\f3\fs20\cf0 
\par This will close the satellite selection dialog box and display the selected satellites.
\par 
\par }
5010
Scribble5010
Select Satellite Apply




Writing



FALSE
6
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}{\f4\fswiss\fprq2 System;}}
{\colortbl\red0\green0\blue0;}
\deflang2057\pard\plain\f3\fs20\cf0 This will  display the selected satellites without closing the selection dialog box\plain\f2\fs20\cf0 
\par 
\par 
\par }
5020
Scribble5020
Select Satellite Cancel




Writing



FALSE
6
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}{\f4\fswiss\fprq2 System;}}
{\colortbl\red0\green0\blue0;}
\deflang2057\pard\plain\f3\fs20\cf0 
\par This will close the satellite selection dialog box and ignore any changes you have made, Any Element updates will be saved.
\par 
\par }
5030
Scribble5030
Satellite Selection clear




Writing



FALSE
6
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}{\f4\fswiss\fprq2 System;}}
{\colortbl\red0\green0\blue0;}
\deflang2057\pard\plain\f3\fs20\cf0 
\par This button will clear all of the selected satellites
\par 
\par }
5040
Scribble5040
Satellite selection details




Writing



FALSE
6
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}{\f4\fswiss\fprq2 System;}}
{\colortbl\red0\green0\blue0;}
\deflang2057\pard\plain\f3\fs20\cf0 
\par This buton will display details of the satellites elements and orbit.
\par 
\par }
5050
Scribble5050
Satellite selection update




Writing



FALSE
6
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}{\f4\fswiss\fprq2 System;}}
{\colortbl\red0\green0\blue0;\red128\green0\blue0;\red0\green128\blue0;}
\deflang2057\pard\plain\f3\fs20\cf0 
\par This button will update the elements via the internet. Please see Updating elemenst via the \plain\f3\fs20\cf2\strike internet\plain\f3\fs20\cf1 \{linkID=37\}\plain\f3\fs20\cf0  for more details
\par 
\par }
5060
Scribble5060
Satellite Selection database




Writing



FALSE
6
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}{\f4\fswiss\fprq2 System;}}
{\colortbl\red0\green0\blue0;}
\deflang2057\pard\plain\f3\fs20\cf0 
\par This drop down list contains all of the element sets available. Selecting a new element set will display details of the satellites in the main selection window and allow them to be selected for display.
\par 
\par }
5070
Scribble5070
Satellite selection main list




Writing



FALSE
6
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}{\f4\fswiss\fprq2 System;}}
{\colortbl\red0\green0\blue0;}
\deflang2057\pard\plain\f3\fs20\cf0 
\par This list displays the satellites available from the selected database of elements. To select a satellite click in the box next to the designator. The list may be sorted by clicking on the column headings
\par 
\par }
5080
Scribble5080
Satellite selection kepsdetails




Writing



FALSE
6
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}{\f4\fswiss\fprq2 System;}}
{\colortbl\red0\green0\blue0;}
\deflang2057\pard\plain\f2\fs20\cf0 
\par This shows the ages of the elemenst in the selected database. A warning will appear if the elemnts are older than the setting specified in the program options.
\par 
\par }
0
0
0
13
1 Overview
2 Overview=Scribble20
2 Acknowledgments=Scribble12
1 Getting Started
2 Setting up the program=Scribble40
2 Getting Started=Scribble30
2 Satellite Selection=Scribble35
2 Main Options=Scribble50
1 Orbital Parameters
2 Keplarian Elements=Scribble60
2 Element Warning=Scribble15
2 Internet Update=Scribble37
2 2 Line Element Sets=Scribble39
6
*InternetLink
16711680
Courier New
0
10
1
....
0
0
0
0
0
0
*ParagraphTitle
0
Arial
0
11
1
B...
0
0
0
0
0
0
*PopupLink
0
Arial
0
8
1
....
0
0
0
0
0
0
*PopupTopicTitle
16711680
Arial
0
10
1
B...
0
0
0
0
0
0
*TopicText
0
Arial
0
10
1
....
0
0
0
0
0
0
*TopicTitle
16711680
Arial
0
16
1
B...
0
0
0
0
0
0
