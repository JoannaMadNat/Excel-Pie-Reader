
#!/usr/local/bin/perl -w
# Change above line to point to your perl binary

use GD::Graph::pie;
use strict;

# Generates a PNG pie chart from 
# Arguments: $worksheet, $imageName, $startRow, $startCol
# Returns: $endRow, $imageSize
sub generatePiePDF() {
    # ARGS
    my $worksheet = $_[0];  # current worksheet
    my $startRow = $_[1];   # Row to start reading from
    my $startCol = $_[2];   # Column to start reading from
    my $imageName = $_[3];  # desired name of PNG image
    # ***

    my @data = ([],[]);

    my $endRow = 0;
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
        
        push(@data[0], $cellVal);

        $cell = $worksheet->get_cell( $row, $startCol + 1 );
        next unless $cell;

        $cellVal = $cell->value();
        push(@data[1], $cellVal);
    }

    my $mygraph = GD::Graph::pie->new(700, 700);
    $mygraph->set(
        '3d'          => 0,
    ) or warn $mygraph->error;


    $mygraph->set_value_font(GD::gdGiantFont, 32);
    $mygraph->set( dclrs => [ qw(marine pink lgreen lpurple purple dpurple cyan dgreen dblue) ] );

    my $myimage = $mygraph->plot(\@data) or die $mygraph->error;

    open my $out, '>', "$imageName.png" or die;
    binmode $out;
    print $out $myimage->png;

    return $endRow; # the row that it stopped reading at
}
