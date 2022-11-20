Password/Passphrase Creator v2.3

This is freeware.  Since security is of the upmost these days,
a tool such as this should assist you in protecting your data.

-----------------------------------------------------------------
Modification History
 3 Dec 2000  2.3  Modified password output display by recalculating
                  the number of words per row.  This changed the file
                  output to calculate for a max line output of 68
                  data characters plus carriage return and linefeed 
                  thus totaling 70 characters output per line.  Added
                  an exert of the output file to this Readme.txt file.
 2 Dec 2000  2.2  Corrected password output display by using MS
                  Flexgrid control and fixing the font to "Fixedsys".
                  Enhanced the print output by readjusting the 
                  length of the print string and the title line.
                  Enhanced the "File Save As" to have a more meaningful
                  data header for future reference.  Remove functionality
                  to be able to cut and paste passwords.  User will
                  have to do this manually.  Ms Flexgrid is not that
                  user friendly. 
30 Nov 2000  2.1  changed the number of special characters by minus
                  one.  Changed the password output display to 
                  increase the number of passwords per line.          
18 Nov 2000  2.0  Rewrote the mixing algorithm, random generation,
                  and the unique pointer module.  Removed obsolete
                  routines and variables.  Changed variables to 
                  meet my naming standards.  Added more documentation
                  to the source code. Increased number of passwords
                  that can be generated to 2500.  Increase at your 
                  own risk.  I did not want to run out of resources.
 1 Oct 2000  1.8  Updated the Seeding of the random number generator
                  and modified the Seed2 routine of clsRndData.cls  
23 Sep 2000  1.7  Changed form colors 
10 Jul 2000  1.6  Found too many duplicate passwords when using
                  a minimum length of 8 characters.  Changed 
                  algorithm on creation by creating one long 
                  string of characters and then breaking them up.
 9 Jul 2000  1.5  Added ascending sort to password display.
                  If 2-25 passwords, I use a Bubble sort, if more
                  than 25 passwords, I use a QuickSort.
29 May 2000  1.4  Changed color scheme between passwords and
                  passphrases.  Replaced some BAS modules with
                  class modules.
15 Feb 2000  1.3  Fixed a bug that miscalculated the position 
                  locations within the passwords/passphrase.
                  Started the modification history.
-----------------------------------------------------------------

Note:    After you compile the program, Pwords.dat and Passphrase.exe
	 must be in the same directory

Make sure you have a reference to:
	Microsoft ActiveX Data Objects 2.5 Library
	Microsoft Data Environment Instance 1.0

The DAT file was created using Microsoft Access 2000.  This is
a MDB file renamed.  I use ADODB to access the data.  There are 
six tables loaded with American English words varying in length 
from three to eight characters in length.  This are approximately
26,700 words at your disposal.  If you want to add your own words 
or your own language to the database, just start Microsoft Access 
2000 and open the PWords.dat file.  Add the new words to the 
appropriate table that corresponds to the length of the new word. 
Remember to make a backup of the database before you start making 
changes.  Better safe than sorry.

If you decide to also use numbers and/or the other printable 
characters on the keyboard, then you have a very formidable 
arsenal at your disposal to build your passwords or passphrases.

********************
Options available:
********************

PASSWORD
O	The first letter with all passwords is alphabetic.  Reason
	is that some systems will not accept a number as the first
	character.
O	Set the length of each password.  3 to 20 characters.
O	Choose from 1 to 2500 passwords at a time.
O	You can determine the number of characters to be numeric 
	and/or special characters
O	Sorted in ascending order.


PASSPHRASE
O	Each word in a passphrase can be 3 to 8 characters in length.
	See the menu options.
O	Choose from 1 to 14 words in a passphrase at a time.


BOTH
O	Select whether to use alphabetic, numeric, special keyboard
	characters, or a mix of all the above.
O	Select if the type of display.  Lowercase, uppercase,
        propercase, or a mix of upper and lowercase
O	You can omit specific special characters by using the menu.
O	The ability to highlight and copy the data to another file
	for future reference.
O	Ability to add or delete data from the database tables.
O	You have the ability to save your data to a file or print it.

************************
Sample file output:

Filename:  Test_8.txt
Created:   Sunday  3 December 2000  7:27 AM
 
Length of each word:  6
Number of passwords:  1,000
----------------------------------------------------------------------
aairrr    aancdf    aawgcp    abarzl    abdgbw    abylka    acdioq
acrkas    adiwmp    aeeivm    aevlsh    aftoqc    afzueq    agikez
ahdmyb    ahklth    ahnhdo    ajwuhw    ajzswg    akdazm    akkchj
alwlcu    aneasu    anpnia    antnsn    aoddbl    aodzmh    apbdvk

************************

This has been tested on Windows 98, ME, NT4, 2000.  Let me know what
you think.

-----------------------------------------------------------------
Written by Kenneth Ives                    kenaso@home.com

All of my routines have been compiled with VB6 Service Pack 4.
There are several locations on the web to obtain these
runtime modules.

This software is FREEWARE.  You may use it as you see fit for 
your own projects but you may not re-sell the original or the 
source code. If you redistribute it, you must include this 
disclaimer and all original copyright notices. 

No warranty expressed or implied is given as to the use of this
program.  Use at your own risk.

If you have any suggestions or questions, I would be happy to
hear from you.
-----------------------------------------------------------------
