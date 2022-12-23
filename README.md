# MS_Excel_Word_scraping_with_VBscript
Scraping and modifing data from Microsoft Excel and Word files with VBscript library functions and CLI tool.


<h2>Two parts of the code.</h2>

<h3>vbs_lib</h3>
This is the function library with classes and functions to call from another vbs file.
<br/><br/>
<b>Funcitonalities:</b>
<ul>
  <li>excel and word scraping</li>
<li>utilities for file I/O operations, cli I/O functionalities, cli menu, and config file functionalities, running windows commands in cmd, calculating SHA hashes</li>
</ul>
  
<h3>office_tool</h3>
A windows command line tool with built in menu to call all the functions in the library on an already opened Excel or Word document.
<br/><br/>
<b>Usage of office tool</b><br/>
Run it with office_tool.bat from windows command line or powershell.
Choose between 'Excel' and 'Word'.
Choose the document from the currently open documents to attach to.
Choose the function to call through the menu.





Disclaimer
The purpose of this project was to write some wrapper code around the MS Word and MS Excel parser functions that Microsoft has been implemented in VBscript - to create an easy to use and convenient interface to handle Word and Excel documents.

I've interrupted the project because of the lack of time. I had to focus on other things.

Lot's of useful functions has not been implemented. And the proper error handling is missing to.

But anyway maybe my code will be useful for someone.

Or someone want to upgrade, improve or extend it :)

