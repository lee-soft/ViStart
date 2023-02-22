# GDIPlusWrapper

This library is a continuation of the vbAccelerator GDIPlusWrapper library for the flat GDI API. My goal was to create a more complete GDI+ interface for my own projects. It seems the author was inspired by Microsoft's GDI+ library and to the best of my ability I've tried to continue implementation in this style where possible. 

Since the official vbAccelerator code is no longer available from their website I have included a link to the virgin package for convenience in this repository. It is unmolested and contains the original files intact as they were from the official vbAccelerator website. Out of respect for the original author it is not packaged with my code.

## Background 

Initially I created my software using the flat GDI API without any libraries and it was quite the monolithic nightmare. After I discovered the effecient GDIPlusWrapper by vbAccelerator I began ammending their library to add the functionality that I required in the fashion of Microsoft's implementation of GDI+ in the products they havn't completely neglected yet. 

## Virgin GDIPlusWrapper

Since the library was built on top of another library (vbAccelerator's GDIPlusWrapper) you will need get the original library first. You can find it a copy of it here [vbAccelerator - GDIPlusWrapper](https://github.com/tannerhelland/vbAccelerator-Archive/blob/master/VB/Code/vbMedia/Using_GDI_Plus/GDIPlus_Helper/GDIPlus_Wrapper.zip)

## Getting Started (To compile)

- Ensure you have Visual Basic 6.0(Service Pack 6) installed
- Clone this repository to directory X
- Extract the official virgin GDIPlusWrapper.zip and mate/merge contents with directory X
- Find all patch files and apply the patches to the official GDIPlusWrapper files (IE: Image.cls.patch apply to Image.cls)
- Open the GDIPlusWrapper Extended VB6 project and enjoy the extended GDIPlusWrapper

## How to apply patches

- The .patch files are "diff files". You should apply them the same way you would any diff file. Gnu have a free patch utility for Windows

## Acknolwedgements

I have been unable to contact the original author of the vbAccelerator GDIPlusWrapper (steve@vbaccelerator.com). Permission to include his library here is pending and until it is approved I will not be able to include it here.
