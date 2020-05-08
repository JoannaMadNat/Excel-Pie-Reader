

#!usr/bin/perl
use strict; use warnings;
use PDF::Create;
use utf8;
use Image::Size;
do './drawPie.pl';

# Generates a PDF file from the given worksheet and name
# Arguments: $worksheet, $outputFilename
sub createPDF {
    # ARGS
    my $worksheet = $_[0];      # current worksheet
    my $outputFileName = $_[1]; # desired PDF file name
    # ***

    # Setup pdf utils
    my $pdf = PDF::Create->new(
        'filename'     => "$outputFileName.pdf",
        'Author'       => 'MoJo JoJo',
        'Title'        => 'PerlDF',
        'CreationDate' => [ localtime ]
    );
    my $root = $pdf->new_page('MediaBox' => $pdf->get_page_size('A4'));    
    my $font = $pdf->font('BaseFont' => 'Helvetica');

    my $page1 = $root->new_page;
    my $toc = $pdf->new_outline('Title' => 'P.1', 'Destination' => $page1); # table of contents

    # Setup writer utils
     my $pageWidth = 592;    # get the page dimentions
    my $pageHeight = 840;

    my ( $row_min, $row_max ) = $worksheet->row_range();
    my ( $col_min, $col_max ) = $worksheet->col_range();
    my $leftMargin = 20; 
    my $topMargin = 50; 
    my $fontSize = 28; 
    my $lineCount = $topMargin;

    my $firstRow = 1;
    # scan data regular point content
    for (my $row = $row_min ; $row <= $row_max ; $row++) {
        for (my $col = $col_min; $col <= $col_max ; $col++ ) {
            my $cell = $worksheet->get_cell( $row, $col );
            next unless $cell;

            my $cellVal = $cell->value();
            next unless length($cellVal) != 0; # skip the empty ones

            if($firstRow) {
                # Write Title
                $cell = $worksheet->get_cell( $row, $col );
                $cellVal = $cell->value();
                $page1->stringl($font, $fontSize, $leftMargin, $pageHeight - $lineCount, $cellVal);
                $lineCount += 10;
                $page1->line($leftMargin, $pageHeight - $lineCount, $pageWidth - $leftMargin, $pageHeight - $lineCount);
                
                $lineCount += $fontSize * 0.7; # add space below paragraph
                $leftMargin += 20;         # indent for pretty look
                $fontSize = 12;            # change font size

                $firstRow = 0;
                next;
            }

            
            if(checkPie($worksheet->get_cell( $row, $col + 1 ))) {
               ($row, my $imgSize) = addPie($worksheet, $pdf, $page1, $leftMargin, $pageHeight - $lineCount, $row, $col);
               $lineCount += $imgSize + $fontSize * 2.5; # padding
            }
            elsif ($cellVal=~ m/^- ?\w*/) {
        
               ($row, my $offset) = addList($worksheet, $pdf, $page1, $leftMargin, $pageHeight - $lineCount - $fontSize, $row, $col, $font, $fontSize);
                $lineCount += $offset * ($fontSize + 10) + $fontSize * 1.5; # padding
            }
            else {
                $page1->stringl($font, $fontSize, $leftMargin, $pageHeight - $lineCount, $cellVal);
                $lineCount += $fontSize + 2;
            }
        }
    }    
   
    # Close the file and write the PDF
    $pdf->close;
}

sub checkPie {
    # ARGS
    my $adjacent = $_[0];
    # ***

    return 0 unless $adjacent;

    my $cellVal = $adjacent->value();
    return 0 unless length($cellVal) != 0;

    return 0 unless ($cellVal =~ m/\d+(\.\d*)?%/); # match ex: 23.00% or 23%

    return 1;
}

# Adds a pie chart to the position and margin on the page given
# Arguments: $worksheet, $pdf, $page, $leftMargin, $pos, $startRow, $startCol
# Returns: $lastRow, imgheight
sub addPie {
    # ARGS
    my $worksheet = $_[0];  # current worksheet
    my $pdf = $_[1];        # current pdf
    my $page = $_[2];       # current page
    my $leftMargin = $_[3];     # desired left margin
    my $pos = $_[4];            # desired position on the page
    my $startRow = $_[5];       # Row on worksheet to read data from
    my $startCol = $_[6];       # Column from worksheet to read data from
    # ***
    
 # include a jpeg image with scaling to 20% size
    my $imageName = "jojo";
    my $lastRow = generatePiePDF($worksheet, $startRow, $startCol, $imageName);
    system("convert $imageName.png $imageName.jpg");


    my $jpg = $pdf->image("$imageName.jpg");
    (my $globe_x, my $globe_y) = imgsize("$imageName.jpg");
    my $scale = 0.4;

    $page->image(
        'image'  => $jpg,
        'xscale' => $scale,
        'yscale' => $scale,
        'xpos'   => $leftMargin,
        'ypos'   => $pos - ($globe_y * $scale)
    );

    system("rm $imageName.png $imageName.jpg");
    return ($lastRow, $globe_y * $scale);
}



# Adds a list based on the a group list items
# Arguments: $worksheet, $pdf, $page, $leftMargin, $pos, $startRow, $startCol, $font, $fontSize
# Returns: $lastRow
sub addList {
    # ARGS
    my $worksheet = $_[0];  # current worksheet
    my $pdf = $_[1];        # current pdf
    my $page = $_[2];       # current page
    my $leftMargin = $_[3];     # desired left margin
    my $pos = $_[4];            # desired position on the page
    my $startRow = $_[5];       # Row on worksheet to read data from
    my $startCol = $_[6];       # Column from worksheet to read data from
    my $font = $_[7];
    my $fontSize = $_[8];
    # ***
    
    my $maxTblCols = 3; # arrange them in a nice table manner
    my $maxTblRows = 1;
    my $tblCounter = 0;
    my $leftCount = $leftMargin;
    my $endRow = 0;


    for my $row ( $startRow .. $startRow + 101 ) { # caps at 101
        my $cell = $worksheet->get_cell( $row, $startCol );
        if(!$cell) { # end loop when encountering empty line
            $endRow = $row;
            last;
        }        

        my $cellVal = $cell->value();
         if(length($cellVal) == 0){ # end loop when encountering empty line
            $endRow = $row;
            last;
        } 
        
        $cellVal =~ m/^-/;
        $page->stringl($font, $fontSize, $leftCount, $pos, "::   " . $');
        $leftCount += ($leftMargin * 4);

        $tblCounter++;
        if($tblCounter == $maxTblCols) {
            $tblCounter = 0;
            $pos -= $fontSize + 10;
            $leftCount = $leftMargin; 
            $maxTblRows ++;
        }
    }

    return ($endRow, $maxTblRows);
}
