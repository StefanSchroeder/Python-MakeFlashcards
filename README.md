![Logo](http://github.com/StefanSchroeder/Python-MakeFlashcards/blob/master/doc/makeflashcards.png?raw=true)


Python-Makeflashcards
=====================

This is a python plugin for Openoffice/Libreoffice to create nice flashcards for printing from wordlists.

It works with Openoffice/Libreoffice 1-3 (makeFlashcards.py) and 
with Libreoffice 4 (makeFlashcards4.py; tested with Libreoffice 4.0.3 in Windows) .


Installation
============

Save the makeFlashcards4.py script into the scripts directory below 
your Libreoffice profile directory. 

If you don't know how to find your Libreoffice user profile, refer to

https://wiki.documentfoundation.org/UserProfile

E.g. on Windows XP:

\Documents and Setting\user name\Application Data\libreoffice\4\user\Scripts\python

Or on Windows 7:

c:\Program Files\LibreOffice 4\share\Scripts\python

Strangely it didn't work to put the script in the user specific script folder;
it must be the system wide script folder.

Quick Usage
===========

Create or load a wordlist with three fields per line, TAB separated.

![Logo](http://github.com/StefanSchroeder/Python-MakeFlashcards/blob/master/doc/wordlist.png?raw=true)

Execute the script: Tools -> Macros -> Run Macro.

![Logo](http://github.com/StefanSchroeder/Python-MakeFlashcards/blob/master/doc/runmacro.png?raw=true)

Open the My Macros fold.

Select the script and click Run.

![Logo](http://github.com/StefanSchroeder/Python-MakeFlashcards/blob/master/doc/runmacro2.png?raw=true)

Result:

![Logo](http://github.com/StefanSchroeder/Python-MakeFlashcards/blob/master/doc/flashcards.png?raw=true)


(Somewhat outdated) Documentation is available here:

http://www.tokonoma.de/flashcards.html

Written by Stefan Schroeder

License GNU GPLv3.

Source has been beautified with pytidy.

http://www.gnu.org/licenses/gpl-3.0.html


