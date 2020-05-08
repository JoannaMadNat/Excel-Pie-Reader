

#!usr/bin/perl
use strict; use warnings;
use PDF::Create;
use utf8;

# Takes a list of formatted lines and draws them in the form of a list
# Arguments: $worksheet, $pdf, $page, $leftMargin, $pos, $startRow, $startCol
# Returns: $lastRow
sub addList() {
     # ARGS
    my $worksheet = $_[0];  # current worksheet
    my $pdf = $_[1];        # current pdf
    my $page = $_[2];       # current page
    my $leftMargin = $_[3];     # desired left margin
    my $pos = $_[4];            # desired position on the page
    my $startRow = $_[5];       # Row on worksheet to read data from
    my $startCol = $_[6];       # Column from worksheet to read data from
    # ***

    for my $row ( $startRow .. $startRow + 101 ) { # caps at 101
        my $cell = $worksheet->get_cell( $row, $startCol );
        if($cell == undef) { # end loop when encountering empty line
            $endRow = $row;
            last;
        }        

        my $cellVal = $cell->value();
        if(length($cellVal) == 0){ # end loop when encountering empty line
            $endRow = $row;
            last;
        } 
        
        $cellVal~= m/^- /;
    }