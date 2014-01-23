SimpleWord2013toPDF

*******************************************************************************
Copyright Glenn Mitchell 2014
Thursday, January 23, 2014

A simple converter that will take a string input of a word document and output
a print optimised PDF keeping fonts, bookmarks and content intact.

Requirements:
Microsoft Word 2013 Interop dll

Example Usage:
var converter = new WordToPdf.Converter(inputFile, outputFile);
converter.Convert();
*******************************************************************************
