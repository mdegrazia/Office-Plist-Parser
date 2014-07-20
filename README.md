Office-Plist-Parser
===================

This script parses the recent documents from the MS Office plist file.
The information parsed includes the access dates, file paths and file names. 
Supported versions of Office are Office 2008 and 2010. Other versions may work, but have not been tested.


####Required Library 
  Install the biplist on Linux/OS X using:

    sudo easy_install biplist
    
  For Windows, if you don't already have it installed, you'll need to grab the easy install utility which is included in the   setup tools from python.org, https://pypi.python.org/pypi/setuptools.  The setup tools will place easy_install.exe into your Python directory in the Scripts folder.   Change into this directory and run:

    easy_install.exe biplist
  
  Or download the biplist library from http://github.com/wooster/biplist and manually install it.

####Usage

    OfficePlistParser.py -f com.microsoft.office.plist -o recentdocs.tsv

####More Information

View the blog post at http://az4n6.blogspot.com/2013/08/ms-office-recent-docs-plist-parser.html for more information


Email Mari > arizona4n6 at gmail dot com for help/questions/bugs
