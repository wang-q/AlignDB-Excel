# NAME

AlignDB::Excel - A simple class to use excel to draw charts.

# DESCRIPTION

`AlignDB::Excel` is a  simple class to use excel to draw charts.

Use Win32::OLE module

# ATTRIBUTES

## excel

isa Excel.Application Object

## workbook

isa Excel Workbook Object

## worksheet\_func

isa Excel WorksheetFunction Object

## infile

Input Excel file name

## outfile

Output Excel file name

## font\_name

Which font to use. Default is "Arial"

## font\_size

Font size. Default is 10

## height

Height of generated charts. Default is 200

## width

Width of generated charts. Default is 320

## max\_ticks

Max tick number in the axes. Default is 6

## replace

Replace texts in titles

# METHODS

## BUILD

Init Excel object and open input file.

The BUILD method is called by Moose::Object::BUILDALL, which is
called by Moose::Object::new. So it is also the constructor
method.

## DEMOLISH

instance destructor
save excel file and close excel object

## sheet\_names

Return an ArrayRef contains all worksheet names in the workbook.

## sheet\_name\_set

Return a Set::Scalar object contains all worksheet names in the workbook.

## draw\_y

Draw xlXYScatterLines chart.

## draw\_2y

Draw xlXYScatterLines chart with 2 Y-axis

## draw\_c

Draw xlColumnClustered chart.

## draw\_LineMarkers

Draw xlLineMarkers chart.

## draw\_dd

Draw a special xlLineMarkers chart, distance-density chart.

## draw\_xy

Draw a special xlXYScatter or xlXYScatterLines chart, in which $last\_row is
determined automatically

## get\_column

put column values to an array

## add\_index\_sheet

See HACK #7 in OReilly.Excel.Hacks.2nd.Edition.

This method should be called after all draw\_xxx methods to avoid confusing
those methods.

## time\_stamp

Add a time stamp to worksheet.

## jc\_correction

Do JC correction on some columns.

# AUTHOR

Qiang Wang &lt;wang-q@outlook.com>

# COPYRIGHT AND LICENSE

This software is copyright (c) 2008 by Qiang Wang.

This is free software; you can redistribute it and/or modify it under
the same terms as the Perl 5 programming language system itself.
