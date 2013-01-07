package AlignDB::Excel;

# ABSTRACT: A simple class to use excel to draw charts.

use Moose;

use Win32::OLE qw(in);    # "with" show conflicting with Moose
use Win32::OLE::Const 'Microsoft Excel';
use Win32::OLE::Variant;
use Win32::OLE::NLS qw(:LOCALE :DATE);
$Win32::OLE::Warn = 3;    # die on errors...

use Set::Light;
use List::MoreUtils qw( all any );
use Path::Class;
use Chart::Math::Axis;
use YAML qw(Dump Load DumpFile LoadFile);

=attr excel

isa Excel.Application Object

=cut

has 'excel' => ( is => 'ro', isa => 'Object' );

=attr workbook

isa Excel Workbook Object

=cut

has 'workbook' => ( is => 'ro', isa => 'Object' );

=attr worksheet_func

isa Excel WorksheetFunction Object

=cut

has 'worksheet_func' => ( is => 'ro', isa => 'Object' );

=attr infile

Input Excel file name

=cut

has 'infile' => ( is => 'ro', isa => 'Str', required => 1 );

=attr outfile

Output Excel file name

=cut

has 'outfile' => ( is => 'ro', isa => 'Str' );

=attr font_name

Which font to use. Default is "Arial"

=cut

has 'font_name' => ( is => 'rw', isa => 'Str', default => sub {'Arial'}, );

=attr font_size

Font size. Default is 10

=cut

has 'font_size' => ( is => 'rw', isa => 'Num', default => sub {10}, );

=attr height

Height of generated charts. Default is 200

=cut

has 'height' => ( is => 'rw', isa => 'Num', default => sub {200}, );

=attr width

Width of generated charts. Default is 320

=cut

has 'width' => ( is => 'rw', isa => 'Num', default => sub {320}, );

=attr max_ticks

Max tick number in the axes. Default is 6

=cut

has 'max_ticks' => ( is => 'rw', isa => 'Int', default => sub {6} );

=attr replace

Replace texts in titles

=cut

has 'replace' => ( is => 'rw', isa => 'HashRef', default => sub { {} } );

=method BUILD

      Usage : $obj->BUILD
    Purpose : Init Excel object and open input file.
    Returns : None
 Parameters : None
     Throws : no exceptions
   Comments : The BUILD method is called by Moose::Object::BUILDALL, which is
            : called by Moose::Object::new. So it is also the constructor
            : method. 
   See Also : n/a

=cut

sub BUILD {
    my $self = shift;

    # Init excel object
    my $excel;
    unless ( $excel = Win32::OLE->new("Excel.Application") ) {
        confess "Cannot init Excel.Application\n";
        return;
    }
    $excel->{DisplayAlerts} = 0;        # Turn off all alert boxes
    $self->{excel}          = $excel;

    # set outfile
    my $infile = file( $self->infile );
    $infile = $infile->absolute->stringify;
    my $outfile;
    unless ($outfile) {
        $outfile = $infile;
        $outfile =~ s/(\.xlsx?)$/.chart$1/;
    }
    $self->{outfile} = $outfile;

    # open Excel file
    my $workbook;
    unless ( $workbook = $excel->Workbooks->Open($infile) ) {
        confess "Cannot open Excel file\n";
        return;
    }
    $self->{workbook} = $workbook;

    # init WorksheetFunction
    my $worksheet_func;
    unless ( $worksheet_func = $excel->WorksheetFunction ) {
        confess "Cannot init WorksheetFunction\n";
        return;
    }
    $self->{worksheet_func} = $worksheet_func;

    return;
}

=method DEMOLISH

instance destructor
save excel file and close excel object

=cut

sub DEMOLISH {
    my $self = shift;

    # get excel objects
    my $excel    = $self->excel;
    my $workbook = $self->workbook;

    my $outfile = $self->outfile;

    # Clean up
    $workbook->SaveAs($outfile);
    $excel->Quit;

    return;
}

sub _replace_text {
    my $self    = shift;
    my $text    = shift;
    my $replace = $self->replace;

    for my $key ( keys %$replace ) {
        my $value = $replace->{$key};
        $text =~ s/$key/$value/gi;
    }

    return $text;
}

=method sheet_names

Return an ArrayRef contains all worksheet names in the workbook.

=cut

sub sheet_names {
    my ($self) = @_;
    my $workbook = $self->workbook;

    my @sheet_names;
    foreach my $sheet ( in $workbook->Worksheets ) {
        push @sheet_names, $sheet->{Name};
    }

    return \@sheet_names;
}

=method sheet_name_set

Return a Set::Light object contains all worksheet names in the workbook.

=cut

sub sheet_name_set {
    my ($self) = @_;

    my $sheet_names_ref = $self->sheet_names();
    my $sheet_name_set  = Set::Light->new(@$sheet_names_ref);

    return $sheet_name_set;
}

=method draw_y

Draw xlXYScatterLines chart.

=cut

sub draw_y {
    my ( $self, $sheet_name, $option ) = @_;

    # get excel objects
    my $excel          = $self->excel;
    my $workbook       = $self->workbook;
    my $worksheet_func = $self->worksheet_func;
    my $sheet_name_set = $self->sheet_name_set;

    my $font_name = $option->{font_name} || $self->font_name;
    my $font_size = $option->{font_size} || $self->font_size;
    my $height    = $option->{Height}    || $self->height;
    my $width     = $option->{Width}     || $self->width;

    # axis titles
    my $x_title = $self->_replace_text( $option->{x_title} );
    my $y_title = $self->_replace_text( $option->{y_title} );

    my $sheet;
    if ( $sheet_name_set->has($sheet_name) ) {
        $sheet = $workbook->Worksheets($sheet_name);
    }
    else {
        return;
    }

    # Range Length satrts at: A1 and finishes at: H3
    my $first_row = $option->{first_row};
    my $last_row  = $option->{last_row};
    my ( $x_column, $y_column ) = ( $option->{x_column}, $option->{y_column} );
    my ($y_last_column) = ( $option->{y_last_column} );
    unless ( defined $y_last_column ) {
        $y_last_column = $y_column;
    }

    my $x_range = $sheet->Range(
        $sheet->Cells( $first_row, $x_column ),
        $sheet->Cells( $last_row,  $x_column )
    );
    my $y_range = $sheet->Range(
        $sheet->Cells( $first_row, $y_column ),
        $sheet->Cells( $last_row,  $y_last_column )
    );
    my $range = $excel->Union( $x_range, $y_range );

    # Set axes' scale
    my $x_max_scale = $option->{x_max_scale};
    my $x_min_scale = $option->{x_min_scale};
    unless ( defined $x_min_scale ) {
        $x_min_scale = 0;
    }
    unless ( defined $x_max_scale ) {
        my $x_scale_unit = $option->{x_scale_unit};
        my $x_min_value  = $worksheet_func->Min($x_range);
        my $x_max_value  = $worksheet_func->Max($x_range);
        $x_min_scale = int( $x_min_value / $x_scale_unit ) * $x_scale_unit;
        $x_max_scale
            = ( int( $x_max_value / $x_scale_unit ) + 1 ) * $x_scale_unit;
    }

    my $y_scale = $self->_find_scale($y_range);

    # Select what type of chart you want
    my $chart = $workbook->Charts->Add;
    $chart->Location( xlLocationAsObject, $sheet_name );

    my $chart_serial = $option->{chart_serial};
    my $chart_object = $sheet->ChartObjects($chart_serial)->{chart};

    # Position, size
    $sheet->ChartObjects($chart_serial)->{Height} = $height;
    $sheet->ChartObjects($chart_serial)->{Width}  = $width;
    $sheet->ChartObjects($chart_serial)->{Top}    = $option->{Top};
    $sheet->ChartObjects($chart_serial)->{Left}   = $option->{Left};

    # ChartType
    $chart_object->{ChartType} = xlXYScatterLines;
    $chart_object->SetSourceData( { Source => $range, PlotBy => xlColumns } );

    # Format
    $chart_object->{HasTitle}                           = 0;
    $chart_object->{HasLegend}                          = 0;
    $chart_object->{PlotArea}->{Interior}->{ColorIndex} = 2;
    $chart_object->{PlotArea}->{Border}->{LineStyle}    = xlLineStyleNone;
    $chart_object->{ChartArea}->{Border}->{LineStyle}   = xlLineStyleNone;
    $chart_object->{ChartArea}->{Font}->{Name}          = $font_name;
    $chart_object->{ChartArea}->{Font}->{Size}          = $font_size;

    # X axis
    $chart_object->Axes(xlCategory)->{Border}->{Weight}           = xlThin;
    $chart_object->Axes(xlCategory)->{HasMajorGridlines}          = 0;
    $chart_object->Axes(xlCategory)->{HasTitle}                   = 1;
    $chart_object->Axes(xlCategory)->{AxisTitle}->{Text}          = $x_title;
    $chart_object->Axes(xlCategory)->{AxisTitle}->{AutoScaleFont} = 1;
    $chart_object->Axes(xlCategory)->{MajorTickMark} = xlTickMarkInside;
    $chart_object->Axes(xlCategory)->{MinimumScale}  = $x_min_scale;
    $chart_object->Axes(xlCategory)->{MaximumScale}  = $x_max_scale;

    # Y axis
    $chart_object->Axes(xlValue)->{Border}->{Weight}           = xlThin;
    $chart_object->Axes(xlValue)->{HasMajorGridlines}          = 0;
    $chart_object->Axes(xlValue)->{HasTitle}                   = 1;
    $chart_object->Axes(xlValue)->{AxisTitle}->{Text}          = $y_title;
    $chart_object->Axes(xlValue)->{AxisTitle}->{AutoScaleFont} = 1;
    $chart_object->Axes(xlValue)->{MajorTickMark} = xlTickMarkInside;
    $chart_object->Axes(xlValue)->{MinimumScale}  = $y_scale->{bottom};
    $chart_object->Axes(xlValue)->{MaximumScale}  = $y_scale->{top};
    $chart_object->Axes(xlValue)->{MajorUnit}     = $y_scale->{unit};

    if ( exists $option->{cross} ) {
        $chart_object->Axes(xlCategory)->{CrossesAt} = $option->{cross};
    }

    return;
}

=method draw_2y

Draw xlXYScatterLines chart with 2 Y-axis

=cut

sub draw_2y {
    my ( $self, $sheet_name, $option ) = @_;

    # get excel objects
    my $excel          = $self->excel;
    my $workbook       = $self->workbook;
    my $worksheet_func = $self->worksheet_func;
    my $sheet_name_set = $self->sheet_name_set;

    my $font_name = $option->{font_name} || $self->font_name;
    my $font_size = $option->{font_size} || $self->font_size;
    my $height    = $option->{Height}    || $self->height;
    my $width     = $option->{Width}     || $self->width;

    # axis titles
    my $x_title  = $self->_replace_text( $option->{x_title} );
    my $y_title  = $self->_replace_text( $option->{y_title} );
    my $y2_title = $self->_replace_text( $option->{y2_title} );

    my $sheet;
    if ( $sheet_name_set->has($sheet_name) ) {
        $sheet = $workbook->Worksheets($sheet_name);
    }
    else {
        return;
    }

    # Range Length satrts at: A1 and finishes at: H3
    my $first_row = $option->{first_row};
    my $last_row  = $option->{last_row};
    my ( $x_column, $y_column ) = ( $option->{x_column}, $option->{y_column} );
    my ($y2_column) = ( $option->{y2_column} );

    my $x_range = $sheet->Range(
        $sheet->Cells( $first_row, $x_column ),
        $sheet->Cells( $last_row,  $x_column )
    );
    my $y_range = $sheet->Range(
        $sheet->Cells( $first_row, $y_column ),
        $sheet->Cells( $last_row,  $y_column )
    );
    my $y2_range = $sheet->Range(
        $sheet->Cells( $first_row, $y2_column ),
        $sheet->Cells( $last_row,  $y2_column )
    );
    my $range = $excel->Union( $x_range, $y_range );

    # Set axes' scale
    my $x_max_scale = $option->{x_max_scale};
    my $x_min_scale = $option->{x_min_scale};
    unless ( defined $x_min_scale ) {
        $x_min_scale = 0;
    }
    unless ( defined $x_max_scale ) {
        my $x_scale_unit = $option->{x_scale_unit};
        my $x_min_value  = $worksheet_func->Min($x_range);
        my $x_max_value  = $worksheet_func->Max($x_range);
        $x_min_scale = int( $x_min_value / $x_scale_unit ) * $x_scale_unit;
        $x_max_scale
            = ( int( $x_max_value / $x_scale_unit ) + 1 ) * $x_scale_unit;
    }

    my $y_scale  = $self->_find_scale($y_range);
    my $y2_scale = $self->_find_scale($y2_range);

    # Select what type of chart you want
    my $chart = $workbook->Charts->Add;
    $chart->Location( xlLocationAsObject, $sheet_name );

    my $chart_serial = $option->{chart_serial};
    my $chart_object = $sheet->ChartObjects($chart_serial)->{chart};

    # Position, size
    $sheet->ChartObjects($chart_serial)->{Height} = $height;
    $sheet->ChartObjects($chart_serial)->{Width}  = $width;
    $sheet->ChartObjects($chart_serial)->{Top}    = $option->{Top};
    $sheet->ChartObjects($chart_serial)->{Left}   = $option->{Left};

    # ChartType
    $chart_object->{ChartType} = xlXYScatterLines;
    $chart_object->SetSourceData( { Source => $range, PlotBy => xlColumns } );

    # Format
    $chart_object->{HasTitle}                           = 0;
    $chart_object->{HasLegend}                          = 0;
    $chart_object->{PlotArea}->{Interior}->{ColorIndex} = 2;
    $chart_object->{PlotArea}->{Border}->{LineStyle}    = xlLineStyleNone;
    $chart_object->{ChartArea}->{Border}->{LineStyle}   = xlLineStyleNone;
    $chart_object->{ChartArea}->{Font}->{Name}          = $font_name;
    $chart_object->{ChartArea}->{Font}->{Size}          = $font_size;

    # X axis
    $chart_object->Axes(xlCategory)->{Border}->{Weight}           = xlThin;
    $chart_object->Axes(xlCategory)->{HasMajorGridlines}          = 0;
    $chart_object->Axes(xlCategory)->{HasTitle}                   = 1;
    $chart_object->Axes(xlCategory)->{AxisTitle}->{Text}          = $x_title;
    $chart_object->Axes(xlCategory)->{AxisTitle}->{AutoScaleFont} = 1;
    $chart_object->Axes(xlCategory)->{MajorTickMark} = xlTickMarkInside;
    $chart_object->Axes(xlCategory)->{MinimumScale}  = $x_min_scale;
    $chart_object->Axes(xlCategory)->{MaximumScale}  = $x_max_scale;

    # Y axis
    $chart_object->Axes(xlValue)->{Border}->{Weight}           = xlThin;
    $chart_object->Axes(xlValue)->{HasMajorGridlines}          = 0;
    $chart_object->Axes(xlValue)->{HasTitle}                   = 1;
    $chart_object->Axes(xlValue)->{AxisTitle}->{Text}          = $y_title;
    $chart_object->Axes(xlValue)->{AxisTitle}->{AutoScaleFont} = 1;
    $chart_object->Axes(xlValue)->{AxisTitle}->{Font}->{Color}
        = RGB( 79, 129, 189 );
    $chart_object->Axes(xlValue)->{MajorTickMark} = xlTickMarkInside;
    $chart_object->Axes(xlValue)->{MinimumScale}  = $y_scale->{bottom};
    $chart_object->Axes(xlValue)->{MaximumScale}  = $y_scale->{top};
    $chart_object->Axes(xlValue)->{MajorUnit}     = $y_scale->{unit};

    # second axis
    $chart_object->SeriesCollection->Add( { Source => $y2_range } );
    $chart_object->SeriesCollection(2)->{AxisGroup}  = xlSecondary;
    $chart_object->SeriesCollection(2)->{MarkerSize} = 5;

    $chart_object->Axes( xlValue, xlSecondary )->{Border}->{Weight}  = xlThin;
    $chart_object->Axes( xlValue, xlSecondary )->{HasMajorGridlines} = 0;
    $chart_object->Axes( xlValue, xlSecondary )->{HasTitle}          = 1;
    $chart_object->Axes( xlValue, xlSecondary )->{AxisTitle}->{Text}
        = $y2_title;
    $chart_object->Axes( xlValue, xlSecondary )->{AxisTitle}->{AutoScaleFont}
        = 1;
    $chart_object->Axes( xlValue, xlSecondary )->{AxisTitle}->{Font}->{Color}
        = RGB( 192, 80, 77 );
    $chart_object->Axes( xlValue, xlSecondary )->{MajorTickMark}
        = xlTickMarkInside;
    $chart_object->Axes( xlValue, xlSecondary )->{MinimumScale}
        = $y2_scale->{bottom};
    $chart_object->Axes( xlValue, xlSecondary )->{MaximumScale}
        = $y2_scale->{top};
    $chart_object->Axes( xlValue, xlSecondary )->{MajorUnit}
        = $y2_scale->{unit};

    return;
}

=method draw_c

Draw xlColumnClustered chart.

=cut

sub draw_c {
    my ( $self, $sheet_name, $option ) = @_;

    # get excel objects
    my $excel          = $self->excel;
    my $workbook       = $self->workbook;
    my $worksheet_func = $self->worksheet_func;
    my $sheet_name_set = $self->sheet_name_set;

    my $sheet;
    if ( $sheet_name_set->has($sheet_name) ) {
        $sheet = $workbook->Worksheets($sheet_name);
    }
    else {
        return;
    }

    # Range Length satrts at: A1 and finishes at: H3
    my $first_row = $option->{first_row};
    my $last_row  = $option->{last_row};
    my ( $x_column, $y_column ) = ( $option->{x_column}, $option->{y_column} );
    my $x_range = $sheet->Range(
        $sheet->Cells( $first_row, $x_column ),
        $sheet->Cells( $last_row,  $x_column )
    );
    my $y_range = $sheet->Range(
        $sheet->Cells( $first_row, $y_column ),
        $sheet->Cells( $last_row,  $y_column )
    );
    my $range = $excel->Union( $x_range, $y_range );

    # There are no  axes' scales to be set

    # Select what type of chart you want
    my $chart = $workbook->Charts->Add;
    $chart->Location( xlLocationAsObject, $sheet_name );

    my $chart_serial = $option->{chart_serial};
    my $chart_object = $sheet->ChartObjects($chart_serial)->{chart};

    # Position, size
    $sheet->ChartObjects($chart_serial)->{Height} = $option->{Height};
    $sheet->ChartObjects($chart_serial)->{Width}  = $option->{Width};
    $sheet->ChartObjects($chart_serial)->{Top}    = $option->{Top};
    $sheet->ChartObjects($chart_serial)->{Left}   = $option->{Left};

    # ChartType
    $chart_object->{ChartType} = xlColumnClustered;
    $chart_object->SetSourceData( { Source => $range, PlotBy => xlColumns } );

    # Format
    $chart_object->{HasTitle}                           = 0;
    $chart_object->{HasLegend}                          = 0;
    $chart_object->{PlotArea}->{Interior}->{ColorIndex} = 2;
    $chart_object->{PlotArea}->{Border}->{LineStyle}    = xlLineStyleNone;
    $chart_object->{ChartArea}->{Font}->{Name}          = "Arial";
    $chart_object->{ChartArea}->{Font}->{Size}          = 12;
    $chart_object->{ChartArea}->{Border}->{LineStyle}   = xlLineStyleNone;

    # axis titles
    my $x_title = $self->_replace_text( $option->{x_title} );
    my $y_title = $self->_replace_text( $option->{y_title} );

    # X axis
    $chart_object->Axes(xlCategory)->{Border}->{Weight}          = xlThin;
    $chart_object->Axes(xlCategory)->{HasMajorGridlines}         = 0;
    $chart_object->Axes(xlCategory)->{HasTitle}                  = 1;
    $chart_object->Axes(xlCategory)->{AxisTitle}->{Text}         = $x_title;
    $chart_object->Axes(xlCategory)->{TickLabels}->{Orientation} = 90;
    $chart_object->Axes(xlCategory)->{MajorTickMark} = xlTickMarkInside;

    # Y axis
    $chart_object->Axes(xlValue)->{Border}->{Weight}  = xlThin;
    $chart_object->Axes(xlValue)->{HasMajorGridlines} = 0;
    $chart_object->Axes(xlValue)->{MinimumScale}      = 0;
    $chart_object->Axes(xlValue)->{HasTitle}          = 1;
    $chart_object->Axes(xlValue)->{AxisTitle}->{Text} = $y_title;
    $chart_object->Axes(xlValue)->{MajorTickMark}     = xlTickMarkInside;

    return;
}

=method draw_LineMarkers

Draw xlLineMarkers chart.

=cut

sub draw_LineMarkers {
    my ( $self, $sheet_name, $option ) = @_;

    # get excel objects
    my $excel          = $self->excel;
    my $workbook       = $self->workbook;
    my $worksheet_func = $self->worksheet_func;
    my $sheet_name_set = $self->sheet_name_set;

    my $font_name = $option->{font_name} || $self->font_name;
    my $font_size = $option->{font_size} || $self->font_size;
    my $height    = $option->{Height}    || $self->height;
    my $width     = $option->{Width}     || $self->width;

    # axis titles
    my $x_title = $self->_replace_text( $option->{x_title} );
    my $y_title = $self->_replace_text( $option->{y_title} );

    my $sheet;
    if ( $sheet_name_set->has($sheet_name) ) {
        $sheet = $workbook->Worksheets($sheet_name);
    }
    else {
        return;
    }

    # Range Length satrts at: A1 and finishes at: H3
    my $first_row = $option->{first_row};
    my $last_row  = $option->{last_row};
    my ( $x_column, $y_column ) = ( $option->{x_column}, $option->{y_column} );
    my ($y_last_column) = ( $option->{y_last_column} );
    unless ( defined $y_last_column ) {
        $y_last_column = $y_column;
    }

    my $x_range = $sheet->Range(
        $sheet->Cells( $first_row, $x_column ),
        $sheet->Cells( $last_row,  $x_column )
    );
    my $y_range = $sheet->Range(
        $sheet->Cells( $first_row, $y_column ),
        $sheet->Cells( $last_row,  $y_last_column )
    );
    my $range = $excel->Union( $x_range, $y_range );

    # Set axes' scale
    my $y_scale = $self->_find_scale($y_range);

    # Set x lable Orientation
    my $x_ori = $option->{x_ori};

    # Select what type of chart you want
    my $chart = $workbook->Charts->Add;
    $chart->Location( xlLocationAsObject, $sheet_name );

    my $chart_serial = $option->{chart_serial};
    my $chart_object = $sheet->ChartObjects($chart_serial)->{chart};

    # Position, size
    $sheet->ChartObjects($chart_serial)->{Height} = $height;
    $sheet->ChartObjects($chart_serial)->{Width}  = $width;
    $sheet->ChartObjects($chart_serial)->{Top}    = $option->{Top};
    $sheet->ChartObjects($chart_serial)->{Left}   = $option->{Left};

    # ChartType
    $chart_object->{ChartType} = xlLineMarkers;
    $chart_object->SetSourceData( { Source => $range, PlotBy => xlColumns } );

    # Format
    $chart_object->{HasTitle}                           = 0;
    $chart_object->{HasLegend}                          = 0;
    $chart_object->{PlotArea}->{Interior}->{ColorIndex} = 2;
    $chart_object->{PlotArea}->{Border}->{LineStyle}    = xlLineStyleNone;
    $chart_object->{ChartArea}->{Border}->{LineStyle}   = xlLineStyleNone;
    $chart_object->{ChartArea}->{Font}->{Name}          = $font_name;
    $chart_object->{ChartArea}->{Font}->{Size}          = $font_size;

    # X axis
    $chart_object->Axes(xlCategory)->{Border}->{Weight}          = xlThin;
    $chart_object->Axes(xlCategory)->{HasMajorGridlines}         = 0;
    $chart_object->Axes(xlCategory)->{HasTitle}                  = 1;
    $chart_object->Axes(xlCategory)->{AxisTitle}->{Text}         = $x_title;
    $chart_object->Axes(xlCategory)->{TickLabels}->{Orientation} = $x_ori;
    $chart_object->Axes(xlCategory)->{MajorTickMark} = xlTickMarkInside;

    # Y axis
    $chart_object->Axes(xlValue)->{Border}->{Weight}           = xlThin;
    $chart_object->Axes(xlValue)->{HasMajorGridlines}          = 0;
    $chart_object->Axes(xlValue)->{HasTitle}                   = 1;
    $chart_object->Axes(xlValue)->{AxisTitle}->{Text}          = $y_title;
    $chart_object->Axes(xlValue)->{AxisTitle}->{AutoScaleFont} = 1;
    $chart_object->Axes(xlValue)->{MajorTickMark} = xlTickMarkInside;
    $chart_object->Axes(xlValue)->{MinimumScale}  = $y_scale->{bottom};
    $chart_object->Axes(xlValue)->{MaximumScale}  = $y_scale->{top};
    $chart_object->Axes(xlValue)->{MajorUnit}     = $y_scale->{unit};

    return;
}

=method draw_dd

Draw a special xlLineMarkers chart, distance-density chart.

=cut

sub draw_dd {
    my ( $self, $sheet_name, $option ) = @_;

    # get excel objects
    my $excel          = $self->excel;
    my $workbook       = $self->workbook;
    my $worksheet_func = $self->worksheet_func;
    my $sheet_name_set = $self->sheet_name_set;

    my $font_name = $option->{font_name} || $self->font_name;
    my $font_size = $option->{font_size} || $self->font_size;
    my $height    = $option->{Height}    || $self->height;
    my $width     = $option->{Width}     || $self->width;

    # axis titles
    my $x_title = $self->_replace_text( $option->{x_title} );
    my $y_title = $self->_replace_text( $option->{y_title} );

    my $sheet;
    if ( $sheet_name_set->has($sheet_name) ) {
        $sheet = $workbook->Worksheets($sheet_name);
    }
    else {
        return;
    }

    my @group_name = @{ $option->{group_name} };
    my ( $paste_top, $paste_left ) = ( 2, 8 );

    my ( $section_top, $section_end, $section_length ) = (
        $option->{section_top},
        $option->{section_end},
        $option->{section_length}
    );
    my ( $series, $category, $value ) = ( 1, 2, 3 );

    my ( $paste_end, $paste_right )
        = ( $paste_top + $section_length - 1,
        $paste_left + scalar @group_name );

    # Copy category
    my $copy_range = $sheet->Range(
        $sheet->Cells(
            $section_top + $section_length * $#group_name, $category
        ),
        $sheet->Cells(
            $section_end + $section_length * $#group_name, $category
        )
    );
    $copy_range->Copy;
    my $paste_range = $sheet->Range( $sheet->Cells( $paste_top, $paste_left ),
        $sheet->Cells( $paste_top + $section_length - 1, $paste_left ) );
    $paste_range->PasteSpecial;

    # Copy value
    for ( my $i = 1; $i <= scalar @group_name; $i++ ) {
        my $copy_range = $sheet->Range(
            $sheet->Cells(
                $section_top + $section_length * ( $i - 1 ), $value
            ),
            $sheet->Cells(
                $section_end + $section_length * ( $i - 1 ), $value
            )
        );
        $copy_range->Copy;
        my $paste_range = $sheet->Range(
            $sheet->Cells( $paste_top,                       $paste_left + $i ),
            $sheet->Cells( $paste_top + $section_length - 1, $paste_left + $i )
        );
        $paste_range->PasteSpecial;
    }

    # Copy series
    for ( my $i = 1; $i <= scalar @group_name; $i++ ) {
        my $copy_range = $sheet->Range(
            $sheet->Cells(
                $section_top + 1 + $section_length * ( $i - 1 ), $series
            ),
            $sheet->Cells(
                $section_top + 1 + $section_length * ( $i - 1 ), $series
            )
        );
        $copy_range->Copy;
        my $paste_range = $sheet->Range(
            $sheet->Cells( $paste_top, $paste_left + $i ),
            $sheet->Cells( $paste_top, $paste_left + $i )
        );
        $paste_range->PasteSpecial;
    }

    # Range Length satrts at: A1 and finishes at: H3
    my $range = $sheet->Range(
        $sheet->Cells( $paste_top, $paste_left ),
        $sheet->Cells( $paste_end, $paste_right )
    );

    my $y_range = $sheet->Range(
        $sheet->Cells( $paste_top + 1, $paste_left + 1 ),
        $sheet->Cells( $paste_end,     $paste_right )
    );

    # Set axes' scale
    my $y_scale = $self->_find_scale($y_range);

    # Select what type of chart you want
    my $chart = $workbook->Charts->Add;
    $chart->Location( xlLocationAsObject, $sheet_name );

    my $chart_serial = $option->{chart_serial};
    my $chart_object = $sheet->ChartObjects($chart_serial)->{chart};

    # Position, size
    $sheet->ChartObjects($chart_serial)->{Height} = $height;
    $sheet->ChartObjects($chart_serial)->{Width}  = $width;
    $sheet->ChartObjects($chart_serial)->{Top}    = $option->{Top};
    $sheet->ChartObjects($chart_serial)->{Left}   = $option->{Left};

    # ChartType
    $chart_object->{ChartType} = xlLineMarkers;
    $chart_object->SetSourceData( { Source => $range, PlotBy => xlColumns } );

    ## series
    ## MarkerSize can be a value from 2 through 72
    ## Excel 2007 has a default value of 7
    #my @styles = (
    #    xlMarkerStyleDiamond,
    #    xlMarkerStyleSquare,
    #    xlMarkerStyleTriangle,
    #    xlMarkerStyleCircle
    #);
    #for (1 .. @group_name) {
    #    my $order = ($_ - 1) % 4 ;
    #    $chart_object->SeriesCollection($_)->{MarkerSize} = 7;
    #    $chart_object->SeriesCollection($_)->{MarkerStyle} = $styles[$order];
    #}

    # Format
    $chart_object->{HasTitle}  = 0;
    $chart_object->{HasLegend} = 0;

    #$chart_object->{Legend}->{Position} = xlLegendPositionTop;
    $chart_object->{PlotArea}->{Interior}->{ColorIndex} = 2;
    $chart_object->{PlotArea}->{Border}->{LineStyle}    = xlLineStyleNone;
    $chart_object->{ChartArea}->{Border}->{LineStyle}   = xlLineStyleNone;
    $chart_object->{ChartArea}->{Font}->{Name}          = $font_name;
    $chart_object->{ChartArea}->{Font}->{Size}          = $font_size;

    my ($x_orientation) = ( $option->{x_orientation} );

    # X axis
    $chart_object->Axes(xlCategory)->{Border}->{Weight}  = xlThin;
    $chart_object->Axes(xlCategory)->{HasMajorGridlines} = 0;
    $chart_object->Axes(xlCategory)->{HasTitle}          = 1;
    $chart_object->Axes(xlCategory)->{AxisTitle}->{Text} = $x_title;
    $chart_object->Axes(xlCategory)->{TickLabels}->{Orientation}
        = $x_orientation;
    $chart_object->Axes(xlCategory)->{MajorTickMark} = xlTickMarkInside;

    # Y axis
    $chart_object->Axes(xlValue)->{Border}->{Weight}           = xlThin;
    $chart_object->Axes(xlValue)->{HasMajorGridlines}          = 0;
    $chart_object->Axes(xlValue)->{HasTitle}                   = 1;
    $chart_object->Axes(xlValue)->{AxisTitle}->{Text}          = $y_title;
    $chart_object->Axes(xlValue)->{AxisTitle}->{AutoScaleFont} = 1;
    $chart_object->Axes(xlValue)->{MajorTickMark} = xlTickMarkInside;
    $chart_object->Axes(xlValue)->{MinimumScale}  = $y_scale->{bottom};
    $chart_object->Axes(xlValue)->{MaximumScale}  = $y_scale->{top};
    $chart_object->Axes(xlValue)->{MajorUnit}     = $y_scale->{unit};

    return;
}

=method draw_xy

Draw a special xlXYScatter or xlXYScatterLines chart, in which $last_row is
determined automatically

=cut

sub draw_xy {
    my ( $self, $sheet_name, $option ) = @_;

    # get excel objects
    my $excel          = $self->excel;
    my $workbook       = $self->workbook;
    my $worksheet_func = $self->worksheet_func;
    my $sheet_name_set = $self->sheet_name_set;

    my $font_name = $option->{font_name} || $self->font_name;
    my $font_size = $option->{font_size} || $self->font_size;
    my $height    = $option->{Height}    || $self->height;
    my $width     = $option->{Width}     || $self->width;

    # chart type
    my $chart_type = $option->{without_line} ? xlXYScatter : xlXYScatterLines;

    # marker size
    my $marker_size = $option->{marker_size};

    # trendline
    my $add_trend = $option->{add_trend};

    # axis titles
    my $x_title = $self->_replace_text( $option->{x_title} );
    my $y_title = $self->_replace_text( $option->{y_title} );

    my $sheet;
    if ( $sheet_name_set->has($sheet_name) ) {
        $sheet = $workbook->Worksheets($sheet_name);
    }
    else {
        return;
    }

    # last row
    my $last_row = $sheet->{UsedRange}->{Rows}->{Count};

    my ( $x_column, $y_column ) = ( $option->{x_column}, $option->{y_column} );
    my $x_range = $sheet->Range(
        $sheet->Cells( 2,         $x_column ),
        $sheet->Cells( $last_row, $x_column )
    );
    my $y_range = $sheet->Range(
        $sheet->Cells( 2,         $y_column ),
        $sheet->Cells( $last_row, $y_column )
    );
    my $range = $excel->Union( $x_range, $y_range );

    # Set axes' scale
    my $x_scale = $self->_find_scale($x_range);
    my $y_scale = $self->_find_scale($y_range);

    # Select what type of chart you want
    my $chart = $workbook->Charts->Add;
    $chart->Location( xlLocationAsObject, $sheet_name );

    my $chart_serial = $option->{chart_serial};
    my $chart_object = $sheet->ChartObjects($chart_serial)->{chart};

    # Position, size
    $sheet->ChartObjects($chart_serial)->{Height} = $height;
    $sheet->ChartObjects($chart_serial)->{Width}  = $width;
    $sheet->ChartObjects($chart_serial)->{Top}    = $option->{Top};
    $sheet->ChartObjects($chart_serial)->{Left}   = $option->{Left};

    # ChartType
    $chart_object->{ChartType} = $chart_type;
    $chart_object->SetSourceData( { Source => $range, PlotBy => xlColumns } );

    # series
    # MarkerSize can be a value from 2 through 72
    # Excel 2007 has a default value of 7
    if ( defined $marker_size ) {
        $chart_object->SeriesCollection(1)->{MarkerSize} = $marker_size;
    }

    # Format
    $chart_object->{HasTitle}                           = 0;
    $chart_object->{HasLegend}                          = 0;
    $chart_object->{PlotArea}->{Interior}->{ColorIndex} = 2;
    $chart_object->{PlotArea}->{Border}->{LineStyle}    = xlLineStyleNone;
    $chart_object->{ChartArea}->{Border}->{LineStyle}   = xlLineStyleNone;
    $chart_object->{ChartArea}->{Font}->{Name}          = $font_name;
    $chart_object->{ChartArea}->{Font}->{Size}          = $font_size;

    # X axis
    $chart_object->Axes(xlCategory)->{Border}->{Weight}           = xlThin;
    $chart_object->Axes(xlCategory)->{HasMajorGridlines}          = 0;
    $chart_object->Axes(xlCategory)->{HasTitle}                   = 1;
    $chart_object->Axes(xlCategory)->{AxisTitle}->{Text}          = $x_title;
    $chart_object->Axes(xlCategory)->{AxisTitle}->{AutoScaleFont} = 1;
    $chart_object->Axes(xlCategory)->{MajorTickMark} = xlTickMarkInside;
    $chart_object->Axes(xlCategory)->{MinimumScale}  = $x_scale->{bottom};
    $chart_object->Axes(xlCategory)->{MaximumScale}  = $x_scale->{top};
    $chart_object->Axes(xlCategory)->{MajorUnit}     = $x_scale->{unit};

    # Y axis
    $chart_object->Axes(xlValue)->{Border}->{Weight}           = xlThin;
    $chart_object->Axes(xlValue)->{HasMajorGridlines}          = 0;
    $chart_object->Axes(xlValue)->{HasTitle}                   = 1;
    $chart_object->Axes(xlValue)->{AxisTitle}->{Text}          = $y_title;
    $chart_object->Axes(xlValue)->{AxisTitle}->{AutoScaleFont} = 1;
    $chart_object->Axes(xlValue)->{MajorTickMark} = xlTickMarkInside;
    $chart_object->Axes(xlValue)->{MinimumScale}  = $y_scale->{bottom};
    $chart_object->Axes(xlValue)->{MaximumScale}  = $y_scale->{top};
    $chart_object->Axes(xlValue)->{MajorUnit}     = $y_scale->{unit};

    if ($add_trend) {
        $chart_object->SeriesCollection(1)->Trendlines->Add(
            {   Type            => xlLinear,
                Name            => "Linear Trend",
                DisplayRSquared => 0,
                DisplayEquation => 0,
            }
        );
        my $trendline = $chart_object->SeriesCollection(1)->Trendlines(1);
        $trendline->{Border}->{ColorIndex} = 5;
        $trendline->{Border}->{Weight}     = xlMedium;
        $trendline->{Border}->{LineStyle}  = xlContinuous;
        $trendline->{Border}->{Weight}     = xlMedium;
    }

    return;
}

sub linear_fit {
    my ( $self, $sheet_name, $option ) = @_;

    # get excel objects
    my $excel          = $self->excel;
    my $workbook       = $self->workbook;
    my $worksheet_func = $self->worksheet_func;
    my $sheet_name_set = $self->sheet_name_set;

    my $sheet;
    if ( $sheet_name_set->has($sheet_name) ) {
        $sheet = $workbook->Worksheets($sheet_name);
    }
    else {
        return;
    }

    # last row
    my $last_row = $sheet->{UsedRange}->{Rows}->{Count};

    my ( $x_column, $y_column ) = ( $option->{x_column}, $option->{y_column} );
    my $x_range = $sheet->Range(
        $sheet->Cells( 2,         $x_column ),
        $sheet->Cells( $last_row, $x_column )
    );
    my $y_range = $sheet->Range(
        $sheet->Cells( 2,         $y_column ),
        $sheet->Cells( $last_row, $y_column )
    );

    my $x = [ $self->_all_in_range($x_range) ];
    my $y = [ $self->_all_in_range($y_range) ];

    #print Dump [$x, $y];

    my ( $r_squared, $p_value, $intercept, $slope ) = _r_lm( $x, $y );

    my $chart_serial = $option->{chart_serial};
    my $gap = 15 * ( $chart_serial - 1 );

    #print Dump [$r_squared, $p_value];

    # write values to cells
    $sheet->Cells( 3 + $gap, 16 )->{Value}        = 'r_squared';
    $sheet->Cells( 3 + $gap, 16 )->{Font}->{Name} = $self->font_name;
    $sheet->Cells( 3 + $gap, 16 )->{Font}->{Size} = $self->font_size;
    $sheet->Cells( 3 + $gap, 17 )->{Value}        = $r_squared;
    $sheet->Cells( 3 + $gap, 17 )->{Font}->{Name} = $self->font_name;
    $sheet->Cells( 3 + $gap, 17 )->{Font}->{Size} = $self->font_size;

    $sheet->Cells( 4 + $gap, 16 )->{Value}        = 'p_value';
    $sheet->Cells( 4 + $gap, 16 )->{Font}->{Name} = $self->font_name;
    $sheet->Cells( 4 + $gap, 16 )->{Font}->{Size} = $self->font_size;
    $sheet->Cells( 4 + $gap, 17 )->{Value}        = $p_value;
    $sheet->Cells( 4 + $gap, 17 )->{Font}->{Name} = $self->font_name;
    $sheet->Cells( 4 + $gap, 17 )->{Font}->{Size} = $self->font_size;

    $sheet->Cells( 5 + $gap, 16 )->{Value}        = 'intercept';
    $sheet->Cells( 5 + $gap, 16 )->{Font}->{Name} = $self->font_name;
    $sheet->Cells( 5 + $gap, 16 )->{Font}->{Size} = $self->font_size;
    $sheet->Cells( 5 + $gap, 17 )->{Value}        = $intercept;
    $sheet->Cells( 5 + $gap, 17 )->{Font}->{Name} = $self->font_name;
    $sheet->Cells( 5 + $gap, 17 )->{Font}->{Size} = $self->font_size;

    $sheet->Cells( 6 + $gap, 16 )->{Value}        = 'slope';
    $sheet->Cells( 6 + $gap, 16 )->{Font}->{Name} = $self->font_name;
    $sheet->Cells( 6 + $gap, 16 )->{Font}->{Size} = $self->font_size;
    $sheet->Cells( 6 + $gap, 17 )->{Value}        = $slope;
    $sheet->Cells( 6 + $gap, 17 )->{Font}->{Name} = $self->font_name;
    $sheet->Cells( 6 + $gap, 17 )->{Font}->{Size} = $self->font_size;

    return;
}

# Fitting Linear Models using R
sub _r_lm {
    my $x = shift;
    my $y = shift;

    confess "Give two array-refs to me\n" if ref $x ne 'ARRAY';
    confess "Give two array-refs to me\n" if ref $y ne 'ARRAY';
    confess "Variable lengths differ\n"   if @$x != @$y;
    return                              if @$x <= 2;

    require Statistics::R;

    # Create a communication bridge with R and start R
    my $R = Statistics::R->new;

    $R->set( 'x', $x );
    $R->set( 'y', $y );
    $R->run(q{ fit = lm(y ~ x) });
    $R->run(q{ r_squared <- summary(fit)$r.squared });
    $R->run(q{ intercept <- summary(fit)$coefficients[1] });
    $R->run(q{ slope <- summary(fit)$coefficients[2] });
    $R->run(
        q{
        lmp <- function (modelobject) {
            if (class(modelobject) != "lm") stop("Not an object of class 'lm'")
            f <- summary(modelobject)$fstatistic
            p <- pf(f[1],f[2],f[3],lower.tail=F)
            attributes(p) <- NULL
            return(p)
        }
    }
    );
    $R->run(q{ p_value <- lmp(fit) });

    my $r_squared = $R->get('r_squared');
    my $p_value   = $R->get('p_value');
    my $intercept = $R->get('intercept');
    my $slope     = $R->get('slope');

    $R->stop;

    return ( $r_squared, $p_value, $intercept, $slope );
}

=method get_column

put column values to an array

=cut

sub get_column {
    my $self       = shift;
    my $sheet_name = shift;
    my $column     = shift;

    # get excel objects
    my $excel          = $self->excel;
    my $workbook       = $self->workbook;
    my $worksheet_func = $self->worksheet_func;
    my $sheet_name_set = $self->sheet_name_set;

    my $sheet;
    if ( $sheet_name_set->has($sheet_name) ) {
        $sheet = $workbook->Worksheets($sheet_name);
    }
    else {
        return;
    }

    # last row
    my $last_row = $sheet->{UsedRange}->{Rows}->{Count};
    my $range = $sheet->Range( $sheet->Cells( 2, $column ),
        $sheet->Cells( $last_row, $column ) );

    my $array_ref = [ $self->_all_in_range($range) ];

    return $array_ref;
}

=method add_index_sheet

See HACK #7 in OReilly.Excel.Hacks.2nd.Edition.

This method should be called after all draw_xxx methods to avoid confusing
those methods.

=cut

sub add_index_sheet {
    my $self = shift;

    # get excel objects
    my $excel          = $self->excel;
    my $workbook       = $self->workbook;
    my $worksheet_func = $self->worksheet_func;

    # create a new worksheet named "INDEX"
    my $sheet_name = "INDEX";
    my $sheet
        = $workbook->Worksheets->Add( { Before => $workbook->Worksheets(1) } )
        or croak Win32::OLE->LastError();
    $sheet->{Name} = $sheet_name;
    $sheet->Cells( 1, 1 )->{Value}        = $sheet_name;
    $sheet->Cells( 1, 1 )->{Name}         = $sheet_name;
    $sheet->Cells( 1, 1 )->{Font}->{Name} = $self->font_name;
    $sheet->Cells( 1, 1 )->{Font}->{Size} = $self->font_size;

    my $i = 1;
    foreach my $wsheet ( in $workbook->Worksheets ) {
        next if $wsheet->{Name} eq $sheet_name;
        $i++;

        # Add Hyperlinks to every sheets
        my $range        = $wsheet->Range("A1");
        my $wsheet_index = "Start" . $wsheet->Index;
        $range->{Name} = $wsheet_index;
        $wsheet->Hyperlinks->Add(
            {   Anchor        => $range,
                Address       => "",
                SubAddress    => "$sheet_name",
                TextToDisplay => "Back to Index",
            }
        );
        $range->{Font}->{Name} = $self->font_name;
        $range->{Font}->{Size} = $self->font_size;

        # Add Hyperlinks to index sheet
        $sheet->Hyperlinks->Add(
            {   Anchor        => $sheet->Cells( $i, 1 ),
                Address       => "",
                SubAddress    => $wsheet_index,
                TextToDisplay => $wsheet->Name,
            }
        );
        $sheet->Cells( $i, 1 )->{Font}->{Name} = $self->font_name;
        $sheet->Cells( $i, 1 )->{Font}->{Size} = $self->font_size;
    }

    # set hyperlink column with large width
    $sheet->Columns(1)->{ColumnWidth} = 30;

    return;
}

=method time_stamp

Add a time stamp to worksheet.

=cut

sub time_stamp {
    my ( $self, $sheet_name ) = @_;

    # get excel objects
    my $excel          = $self->excel;
    my $workbook       = $self->workbook;
    my $worksheet_func = $self->worksheet_func;
    my $sheet_name_set = $self->sheet_name_set;

    my $sheet;
    if ( $sheet_name_set->has($sheet_name) ) {
        $sheet = $workbook->Worksheets($sheet_name);
    }
    else {
        return;
    }

    # last row
    my $last_row = $sheet->{UsedRange}->{Rows}->{Count};

    my $now = scalar localtime;
    $sheet->Cells( $last_row + 5, 1 )->{Value}        = $now;
    $sheet->Cells( $last_row + 5, 1 )->{Font}->{Bold} = 1;
    $sheet->Cells( $last_row + 5, 1 )->{Font}->{Name} = $self->font_name;
    $sheet->Cells( $last_row + 5, 1 )->{Font}->{Size} = $self->font_size;

    return;
}

=method jc_correction

Do JC correction on some columns.

=cut

sub jc_correction {
    my ($self) = @_;

    # get excel objects
    my $excel          = $self->excel;
    my $workbook       = $self->workbook;
    my $worksheet_func = $self->worksheet_func;

    my @headers = qw{
        AVG_pi AVG_d_indel AVG_d_noindel AVG_d_complex
        AVG_d_ir AVG_d_nr AVG_d_tr AVG_d_qr
    };

    foreach my $sheet ( in $workbook->Worksheets ) {

        my $last_col = $sheet->UsedRange->Find(
            {   What            => "*",
                SearchDirection => xlPrevious,
                SearchOrder     => xlByColumns
            }
        )->{Column};

        my $last_row = $sheet->UsedRange->Find(
            {   What            => "*",
                SearchDirection => xlPrevious,
                SearchOrder     => xlByRows
            }
        )->{Row};

        my $header_range = $sheet->Range( $sheet->Cells( 1, 1 ),
            $sheet->Cells( 1, $last_col ) );

        my @jc_columns;
        foreach my $header_cell ( in $header_range) {
            my $header_value = $header_cell->Value;
            if ( $header_value and any { $_ eq $header_value } @headers ) {
                push @jc_columns, $header_cell->Column;
                my $corrected_header = $header_value;
                $corrected_header =~ s/AVG_/AVG_jc_/;
                $header_cell->{Value} = $corrected_header;
                $header_cell->Font->{Italic} = "True";
            }
        }

        foreach (@jc_columns) {
            my $jc_range = $sheet->Range( $sheet->Cells( 2, $_ ),
                $sheet->Cells( $last_row, $_ ) );
            foreach my $cur_cell ( in $jc_range) {
                my $cur_value = $cur_cell->Value;
                if ( defined $cur_value ) {
                    $cur_cell->{Value} = $self->_jc_correct($cur_value);
                    $cur_cell->Font->{Italic} = "True";
                }
            }
        }
    }

    return;
}

sub _jc_correct {
    my $self  = shift;
    my $pi    = shift;
    my $jc_pi = -0.75 * log( 1 - ( 4.0 / 3.0 ) * $pi );
    return $jc_pi;
}

sub _all_in_range {
    my $self  = shift;
    my $range = shift;

    my @values;
    for my $cur_cell ( in $range) {
        my $cur_value = $cur_cell->Value;
        if ( defined $cur_value ) {
            push @values, $cur_value;
        }
    }

    return @values;
}

sub _find_scale {
    my $self  = shift;
    my $range = shift;

    my $axis    = Chart::Math::Axis->new;
    my @dataset = $self->_all_in_range($range);

    if (@dataset) {
        $axis->add_data(@dataset);
        $axis->set_maximum_intervals( $self->max_ticks );

        return {
            top    => $axis->top,
            bottom => $axis->bottom,
            unit   => $axis->interval_size,
        };
    }
    else {
        return {
            top    => 1,
            bottom => 0,
            unit   => 0.2,
        };
    }
}

sub RGB {
    my ( $red, $green, $blue ) = @_;
    return $red | ( $green << 8 ) | ( $blue << 16 );
}

1;

__END__

=head1 DESCRIPTION

C<AlignDB::Excel> is a  simple class to use excel to draw charts.

Use Win32::OLE module

=cut
