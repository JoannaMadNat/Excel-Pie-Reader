#!usr/bin/perl
use strict; use warnings;
do './readEX.pl';
do './createPDF.pl';

my $argc = $#ARGV + 1;
if ($argc < 1) {
    print "Usage: main.pl file.xlsx [pdf_name]\n";
    exit;
}

my $fileName = $ARGV[0];
if(! ($fileName =~ m/\.xlsx$/)) {
    print "Wrong file extention\n";
    exit;
}

my $outfileName = "output";
if($argc == 2) {
    $outfileName = $ARGV[1];
    if(($outfileName =~ m/\.pdf$/)) {
        $outfileName = $`;
    }
}

my $worksheet = readEX($fileName);

my $PDFName = $outfileName;
createPDF($worksheet, $PDFName);

system("evince $PDFName.pdf");

# system("rm $PDFName.pdf"); # use for cleaning