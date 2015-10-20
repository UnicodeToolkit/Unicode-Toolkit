#Unicode Toolkit
<p align="center"><img src="https://raw.githubusercontent.com/UnicodeToolkit/Unicode-Toolkit/master/screenshot.jpg" /></p>

This is a small offline tool for Windows that allows you to:

1. Look up the code point for any character (you enter á, it returns 225)
2. Look up the character for any code point (you enter 225, it returns á)
3. Replace all diacritics with their closest ASCII equivalents (a.k.a. accent folding)

<b>Character to code</b><br>
This allows you to type in any character and see it's decimal, hexadecimal and octal code point. 
For example, if you type in á it will return 225 (decimal), 00E1 (hexadecimal), and 0341 (octal). 

<b>Code to character</b><br>
This allows you to type in any Unicode code point and see its respective character. 
For example, if you type in 225 it will display the character á.
It will also tell you the character's name, unicode block, and plane.

<b>Diacritic remover</b><br>
This allows you to remove all diacritics from a block of text and replace them with their closest ASCII equivalents. 
For example, if you type in:

`Ｔẖȉṥ ïṧ ā ẗểṧț`

it will return:

`This is a test`

This program was written in Visual Basic 6 which means it should work on any modern version of Windows but just in case you need to download the VB6 runtimes you can do so from here:

<a href="https://www.microsoft.com/en-us/download/details.aspx?id=24417">https://www.microsoft.com/en-us/download/details.aspx?id=24417</a>
