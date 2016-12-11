# Blinx

Blinx is a Microsoft Word add-in that converts Bible references within Word into Bible links: Have a look at [this example Word document](docs/Example article with Bible links.doc) or even [this pdf](docs/Example article with Bible links.pdf). These Blinks contain the passage (stored as an endnote in the Word document), which becomes visible as a tooltip when the mouse pointer is hovering over it. They are also hyperlinks to an online Bible ("ctrl-left click" on the link). On computers where the Blinx add-on and [BibleWorks](http://www.bibleworks.com) are installed, a "right-click" on the link opens it directly in BibleWorks.

### Requirements<sup>[**[1]**](#_ftn1)</sup>
- Windows (Windows XP+)
- Microsoft Word (2003+)
- Internet Explorer 8+ with an active internet connection or alternatively BibleWorks (7+)

### Installation
- Close Microsoft Word and Outlook.
- Copy "Blinx.dot" into your Microsoft Word start-up folder:<sup>[[2]](#_ftn2)</sup>
  * Possibly: C:\Program Files\Microsoft Office\OfficeXX\Startup
  * or: C:\Documents and Settings\YOUR_USER_NAME\Application Data\Microsoft\Word\Startup
- The add-in will now be loaded automatically when Word is started.

### Usage
- **Important: The execution of Blinx can always be interrupted by pressing "Ctrl-Break"**
- The add-in contains 5 functions that can either be accessed through the Blinx toolbar (in the Add-Ins tab in Word 2007+) or through keyboard shortcuts:
  * ![Image](assets/clip_image002.jpg) **Create Blink (Alt-B)**: Converts all Bible references within the selection or the closest one to the left of the cursor into a pure Bible link without passage.
  * ![Image](assets/clip_image003.jpg) **Create Blink & tooltip (Alt-B)**: Converts all Bible references within the selection or the closest one to the left of the cursor into a Bible link with the passage as tooltip.
  * ![Image](assets/clip_image004.jpg) **Create Blink & insert text (Alt-Shift-B)**: Converts all Bible references within the selection or the closest one to the left of the cursor into a Bible link and inserts the passage into the text.
  * ![Image](assets/clip_image005.jpg) **Unlink Blinks and Hyperlinks (Alt-U)**: Converts all Blinks or hyperlinks within the selection or the closest one to the left of the cursor into normal text.
  * ![Image](assets/clip_image006.jpg) **Open Blinx Options dialog (Ctrl-Alt-Shift-B)**:
    - Bible translation: ESV, NIV, NASB, KJV, ...
    - Online Bible for hyperlinks
    - Length of passage tooltips
    - Reset Blinx (which can solve a few issues)
    - Abbreviations for all Bible books, which can be edited with a double-click.<sup>[[3]](#_ftn3)</sup>

### Copyright
- The current version of Blinx obtains Bible passages either from BibleWorks (if installed) or otherwise from [www.biblegateway.com](http://www.biblegateway.com/) for this initial proof of concept.
- All modern Bible versions are copyrighted. See [www.biblegateway.com/version](http://www.biblegateway.com/version) for copyright regulations of various Bible versions. For extensive quotes in a publication (usually if over 200 verses or more than 10% of a biblical book), a permission in writing needs to be obtained from the appropriate copyright owner.
- Biblegateway.com and BibleWorks do not _appear_ to require any additional copyright notice for quotations acquired through their website/software.

### Known issues
- The reference links contain a special Unicode space between the book name and the chapter number. If it is not displayed as a space, choose an appropriate font (e.g. Times New Roman, Arial, Tahoma, Calibri ...)
- A bug in BibleWorks means that every creation of a Blink will reset the display versions to what they were at the start-up of BibleWorks. If you restart (or just close) BibleWorks, the current display versions will be the new default versions for every following Blink creation.
- Running Blinx might reset some formatting options of a previous search/replace in Word.
- The initial start-up of BibleWorks or Internet Explorer (hidden) can take up to 20 sec.
- BibleGateway limits passage lookups to about 5500 words.
- Further issues are mentioned in [Issues.doc](docs/Issues.doc).

Let me know about bugs or improvements that would be useful: [blinx.add.in@gmail.com](mailto:blinx.add.in@gmail.com)

### Roadmap
- For the present Word add-on, there are 2 remaining objectives:
  - Move to a public API like https://bibles.org/pages/api to obtain the online passage.
  - Allow to switch between English and German Bibles and Bible references through the options dialog. At the moment, such a switch needs to happen in the code and has a couple of bugs.
- A complete redesign of the core functionality of the plugin into a cross-platform library that could also be made available through a public API would be desirable.
- A plugin system could be used to add any data source for retrieving Scripture passages (e.g., BibleWorks, Logos, theWord, free online Bibles & Bible APIs).
- Custom add-ons for Word, Open Office Writer, Adobe Acrobat, Google Docs could then be added to creating Bible links on the fly.
- (The [example pdf](docs/Example article with Bible links.pdf) linked above was generated with an alpha version of such an Adobe Acrobat plugin.)
- I'm currently looking for other developers to join the project before embarking on it.


* * *

<a name="_ftn1"></a>[1]
The add-in should also work with some earlier versions of these applications.

<a name="_ftn2"></a>[2]
The location of the start-up folder can also be checked and changed in Word:
- Word 2003: "Tools | Options | File Locations | Startup"
- Word 2007: "Office Button | Word Options | Advanced tab | File Locations"
- Word 2010: "File | Options | Advanced tab | File Locations"

<a name="_ftn3"></a>[3]
Blinx will only recognize the abbreviations in this list (allowing for variations like Roman numerals instead of 1/2/3 before book names, additional full stops, additional or fewer spaces…) Chapter-verse separator can be either ":" or ".".  The book name must be separated from the rest of the reference by at least 1 space or 1 full stop (e.g. "Jn 3:16" or "Jn.3:16").
