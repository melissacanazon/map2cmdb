# map2cmdb  ***Bare with me, I just started learning python 2 weeks ago, so my code is not pretty, but it works!***


Couple of scripts I am working on to crawl through a Qualys Map and a CMDB spreadsheet, and compare entries for asset management.  It will eventually connect to the Qualys API, create new tags (if need be), and add tags to assets based on owner in the CMDB.  It also creates a list of assets that are not in the CMDB (or are missing info) to turn over to asset management.

* it does not work perfectly, such as when comparing Ip 2 IP it will match 255.255.255.13 with 255.255.255.139 *
* 

It will eventually connect to the Qualys API, create new tags (if need be), and add tags to assets based on owner in the CMDB

1.	Open Asset Inventory CSV, save to Documents as ‘assetTest.xlsx’
1.1.	Del unnecessary columns so that the columns match:
1.1.1.	 ‘A’ is Asset ID, ‘B’ is CIShort Name, ‘C’ is CIAlias Name, ‘D’ is CIDescription, ‘E’ is IPAddress, ‘F’ is Owner
1.2.	Rename ‘Sheet1’ to ‘ASSETS’
1.3.	save
2.	Open the Qualys Maps csv, save to Documents as ‘maps.xlsx’ with ASCII encoding
2.1.	Del rows 1-7, so that headers are in row 1
2.1.1.	‘A’ is IPHost and ‘B’ is DNSHost
2.2.	Rename ‘Sheet1’ to ‘IPs’
2.3.	Add two more sheets named ‘NoIPMatch’ and ‘Check’
2.4.	Save
3.	Make sure that there is an excel book in the same working directory for outputting data to.
3.1.	The scripts call for the book ‘outTest.xlsx’ , it is probably easiest to save the book as this.  If you change the name of this book (or any of the books or sheet names), than you will have to modify the code to point to the proper book/sheet.
4.	Make sure all scripts and excel books are in the same working directory.
5.	Run map2assets.py  


**** also, the reason for 3 scripts is because each does a different function that I will sometimes use individually.  If you want to do the same, then comment out the last few lines that call for the next script to open ****
