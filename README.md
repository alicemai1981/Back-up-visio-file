# Back-up-visio-file
This code creates multiple backup versions for Microsoft visio file(currently up to 10 version).  Unlike Excel and Word, Visio does not provide the file dialog box functionality in VBA. 
To allow user to pick up the folder where the back up versions will be saved, this code call the Windows API SHBrowseForFolderA to create an old style Windows file dailog box. 
the limitation of this old style file dailog is that it may not be able to show a directory on network even with setting all the parameters. However user can type or paste the network 
directory into the dialog box to access it. 

for learning purpose, some code to use the SHGetPathFromIDListA API was also include but not called by the main Sub.SHGetPathFromIDListA
