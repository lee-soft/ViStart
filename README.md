Since it is becoming increasingly more difficult to find the libraries required to compile ViStart I have attached them here. Find "All TLBs" to grab them all in one convenient package. 

## Libraries
- [Windows Unicode API TypeLib](https://github.com/badcodes/vb6/blob/master/%5BInclude%5D/TypeLib/winu.tlb) - Windows API, stores all the API declarations
- [dseaman@uol.com.br GDI+ Type Library 1.05](http://www.vbaccelerator.com/home/VB/Type_Libraries/GDIPlus_Type_Library/article.asp)
- [All TLBs](https://lee-soft.com/bin/TLB.zip)

## Installation

- Extract ViStart zip contents to any directory
- Copy the "Skins" directory to %AppData%/ViStart (create the directory if it doesn't already exist) and rename it "_skins". If done correctly the following directory should exist %appdata%\ViStart\_skins\Windows 7 Start Menu

## Setting up the development environment

- Clone the repo
- Ensure you have installed:
    1) Microsoft Visual Basic 6.0 or Visual Studio 6.0
    2) [Service Pack 6](https://download.cnet.com/Service-Pack-6-for-Visual-Basic-6-0/3000-2206_4-10726545.html)
    3) [Service Pack 6 Security Rollup Update KB3096896](https://www.microsoft.com/en-us/download/details.aspx?id=50722)
- Grab the SHELLLNK, WinU and GDIPlus TLBs - extract the TLBs and re-add as a reference to the project
- Compile and enjoy

## Acknowledgements

I have been unable to contact the original author of the vbAccelerator GDIPlusWrapper (steve@vbaccelerator.com). Permission to include his library here is pending and until it is approved I will not be able to include it here.
