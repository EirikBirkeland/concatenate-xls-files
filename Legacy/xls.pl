#! /usr/local/bin/perl -w
use strict;
use Spreadsheet::ParseExcel::Simple;
use Spreadsheet::WriteExcel;
 
my @excelfiles = ();
my @data = ();
my $sheetnameinit = 'Market Name';
my $excel_directory = '.';
my $AllMarket = 'AllMarket.xls';
opendir(DIR, $excel_directory) or die "couldn't open $excel_directory: $!\n";
@excelfiles = grep { $_ =~ /\.xls/ } readdir DIR;
closedir DIR;
 
#################################################
## Read files in the directory into an array                      ##
#################################################
for my $f (@excelfiles) 
{
  my $excel = Spreadsheet::ParseExcel::Simple->read($f);
  foreach my $sheet ($excel->sheets) 
  {
    next if $sheet->{'sheet'}{'Name'} !~ /$sheetnameinit/;
    {
      while ($sheet->has_data)
      { 
       my @input = $sheet->next_row;
       push @data , join("\t",@input)."\n";
      }    
    }
  }
}
 
######################################################
## Write the array into the XLS file                                         ##
######################################################
my $workbook = Spreadsheet::WriteExcel->new("$AllMarket");
my $worksheet = $workbook->add_worksheet();
my $bold = $workbook->add_format();
$bold->set_bold();
my $firstline = 1;
my ($x,$y) = (1,0);
foreach my $line (@data)
{
  chomp $line;
  if ($firstline eq 1) # Header Lines
  {
    $firstline++;
    $worksheet->write( 0, 0, "Market", $bold );
    $worksheet->write( 0, 1, "City", $bold );
    $worksheet->write( 0, 2, "States", $bold );
  }
  elsif ($line =~ /Texas/ ) # Check for lines contains Texas only
  {
    $firstline++;
    my @formatline = split( '\t' , $line );
    foreach my $cell (@formatline)
    {
      $worksheet->write($x, $y++, $cell);
    }
    $x++;$y=0;
  }
}
$workbook->close();
