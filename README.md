Office-Plist-Parser
===================

This script parses the recent documents from the MS Office plist file.
The information parsed includes the access dates, file paths and file names. 
Supported versions of Office are Office 2008 and 2010. Other versions may work, but have not been tested.


Required Library: biplist from http://github.com/wooster/biplist
use easy install : sudo easy_install biplist


Usage:

OfficePlistParser.py -f com.microsoft.office.plist -o recentdocs.tsv


More Information:
View the blog post at http://az4n6.blogspot.com/2013/08/ms-office-recent-docs-plist-parser.html for more information


Email Mari > arizona4n6 at gmail dot com for help/questions/bugs
