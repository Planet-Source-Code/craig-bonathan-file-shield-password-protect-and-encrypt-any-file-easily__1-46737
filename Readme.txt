File Shield
===========

File Shield is a file protector application. It will encrypt any file and encose it inside another application (lock.exe), where it can then be decrypted and extracted with a correct password. If necessary, the file can be included with its own unique key, which will identify whether the protected file has been altered. Another useful feature is the expiration date system. You may choose to set a date when the program will no longer extract the file.

Encryption depends on the password. The password is not included in the executable at all. When a password is entered, a unique key is created for it. You cannot reverse a unique key encryption as it is one-way. This means that if you enter the wrong password, its unique key will not match the one that is in the file, so it will not extract. Since the file is encrypted with the password, only the same password will decrypt it. When the file UID generation is turned off, and on the VERY unlikely event that two passwords (one wrong, and one right) create the same unique key, the file will extract, but the data will be incorrect if the wrong password is used. When the UID generation is turned on, even a password with the same unique key will not extract the file because it would end up with an 'Extraction Error.' The file UID generation does not allow modified files to be extracted, which is definitely an advantage.

Any file can be protected with File Shield. When the protected executable is ran and the correct password is entered, it will extract the file, run it (with its associated application if it is a document), and then delete it as soon as the program that opened the file closes.


File Shield's Features
======================

> Very easy to use and fast encryption for ANY file
> Each executable can have its own title, specified by you
> Programs can be executed with a command line if necessary
> Optional File Unique IDs prevent the file being changed (at the cost of creation and extraction speed)
> Optional Expiration dates stop users from extracting the file after a certain date
> Log file creation
> Uses the CommonDialog control to make file location simple


Compiling the Source Code
=========================

When compiling the code, you must ensure that the size of Lock.exe is equal to the constant value of SEFSize in both Protection.bas and Protection2.bas modules. To get the correct value, compile Lock.exe (the FileShield_Lock project) and check its size (in bytes). Copy this value over to the modules, and then recompile both projects. You must do this every time you wish to change the FileShield_Lock project.

Unless you have changed the code, the values already there should be correct. The size seems to be rounded upwards to the nearest hard drive sector boundary, so the program size may vary when compiled on different computers. To check if you have the value correct, run Lock.exe, and it should say 'File data not available.'

Note that this project has a group file, so opening that is the ideal option, instead of the two project files.


Contact
=======

You may contact me at craigthesnowman@yahoo.co.uk, or at cb3software@fire-bug.co.uk


Legal Copyright
===============

Copyright of File Shield is owned by Craig Bonathan. You may not use this source code for any kind of commercial purpose. If you do not understand, or do not agree to these terms, you may not use this code.


Final Words
===========

I have worked very hard on this project, so please obey the copyright and don't forget to comment.
Thankyou for downloading this, and I hope that it proves to be very useful, even if you don't actually want the source code.

Craig Bonathan
CB3 Software