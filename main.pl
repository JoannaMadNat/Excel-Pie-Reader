#!usr/bin/perl
use strict; use warnings;
do './readEX.pl';
do './createPDF.pl';

my $worksheet = readEX("thoughts.xlsx");

my $PDFName = "jo";
createPDF($worksheet, $PDFName);

system("evince $PDFName.pdf");
system("rm $PDFName.pdf");