# SDS2ASM
A Python Project to convert Microsoft SDS Classic format to Apple School Manager format

Our school uses an LMS product called SEQTA. It features a report to export all the student data in Microsoft School Data Sync format. However, we need to upload data to Apple School Manager. The data is similar - several CSV files in a single ZIP, but just different enough to be incompatible.

That's where SDS2ASM comes in. Feed it the Microsoft SDS formatted zip file, and it will create a new zip file with the appended "_asm" moniker in the same directory.

As the command line usage reminds, only 1 parameter is required... the input file.
There are two switches
-h for help
-q for quiet, which will overwrite an older zip file without prompting.
