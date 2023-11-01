# VBA.Compress

## Compression and decompression methods for VBA

Version 1.3.0

*(c) Gustav Brock, Cactus Data ApS, CPH*

### Zip, Cabinet, and Archive Folders
Set of VBA functions to:

- *zip* and *unzip* zip files and folders
- *compress* and *decompress* cab (cabinet) files and folders
- *tar* and *untar* archive folders

for both 32- and 64-bit. No third-party tools used, only a single code module.

The main goal for the functions is not to offer all sorts of fancy compression and expanding methods and options but, with code, to closely mimic what you manually can do with Windows Explorer.

However, while Windows Explorer's single option for compressing files or folders is to right-click and select _Compress to zip file ..._, included here are functions not only to create zip files, but also to create cab, tar, and tgz files.

A secondary goal is to make the call of the functions as simple and possible. As a result, any operation can be performed with a single line of code with only a few (all optional) arguments.

### Zip Folders

![Zip Folders][zip folders]

Windows Explorer let you handle zip folders nearly as any other folder: Copy, move, change, and delete, etc.

In VBA, you can also handle normal files and folders, but to *zip* folders take a little more - and that you'll find here.

### Cabinet Folders

![Cabinet Folders][cab folders]

Windows Explorer lets you open cabinet (cab) files like any other folder, though for reading only. 

In VBA, you can easily handle normal files and folders, but opening and, indeed, creating *cabinet* files take a lot more - and that you'll find here. 

### Archive Folders

![Tar Folders][tar folders]

Windows Explorer (of _Windows 11_ or later) lets you open archive folders (tar and tgz) files like any other folder, though for reading only. 

In VBA, you can easily handle normal files and folders, but opening and, indeed, creating *archive* files take a little more - and that you'll find here, though _Windows 10_ or later is required. 

<hr>

### Demos

Demos (to download) for _Microsoft Access and Excel_ are located in the [demos][demos folder] folder.

### Documentation
Top level documentation generated by [MZ-Tools][mz tools] is included for [Microsoft Access and Excel][code documentation].

Detailed documentation is included as in-line comments. 

Full documentation can be found here:

![EE Logo][ee logo]

[Zip and unzip files and folders with VBA the Windows Explorer way][ee zip files]

[Handle cabinet files and folders with VBA the Windows Explorer way][ee cab files]

[Handle Archive Files and Folders With VBA the Windows Explorer Way][ee tar files]

<hr>

*If you wish to support my work or need extended support or advice, feel free to:*

<p>

![Buy Me a Coffee][buy me a coffee]

<hr>

[zip folders]: images/GitZip11.png
[cab folders]: images/GitCab11.png
[tar folders]: images/GitTar11.png
[ee logo]: images/EE%20Logo.png
[buy me a coffee]: images/BuyMeACoffee.png
[demos folder]: ./demos
[ee zip files]: https://www.experts-exchange.com/articles/31130/Zip-and-unzip-files-and-folders-with-VBA-the-Windows-Explorer-way.html?preview=yvSy86kgNss%3D
[ee cab files]: https://www.experts-exchange.com/articles/31144/Handle-cabinet-files-and-folders-with-VBA-the-Windows-Explorer-way.html?preview=i8Bvq8gkxiA%3D
[ee tar files]: https://www.experts-exchange.com/articles/35655/Handle-Archive-Files-and-Folders-With-VBA-the-Windows-Explorer-Way.html?preview=546Zf6%2BK49U%3D

[mz tools]: https://www.mztools.com/
[code documentation]: https://htmlpreview.github.io?https://github.com/GustavBrock/VBA.Compress/blob/master/documentation/Compress.htm

