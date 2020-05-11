#!usr/bin/perl
use strict; use warnings;
use Spreadsheet::ParseXLSX;
binmode STDOUT, ":utf8";

# reads a XLSX file and returns the first worsheet of the workbook
# arguments: $fileName
# returns:   $worksheet
sub readEX {
    # ARGS
    my $fileName = $_[0];
    # ***

    my $parser = Spreadsheet::ParseXLSX->new;
    my $workbook = $parser->parse($fileName);
    my $print_areas = $workbook->get_print_areas();
    my @worksheets = $workbook->worksheets();
    my $worksheet = $worksheets[0];

    # Let's read the Excel sheet
    my ( $row_min, $row_max ) = $worksheet->row_range();
    my ( $col_min, $col_max ) = $worksheet->col_range();

    for my $row ( $row_min  .. $row_max ) {
        for my $col ( $col_min .. $col_max ) {

            my $cell = $worksheet->get_cell( $row, $col );
            next unless $cell;

            my $cellVal = $cell->value();
            next unless length($cellVal) != 0; # skip the empty ones

            print "($row, $col) : $cellVal\n";
        }
    }


    return $worksheet;
}
