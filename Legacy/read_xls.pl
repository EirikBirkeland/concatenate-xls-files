#!/usr/bin/perl -w

    use strict;
    use Spreadsheet::ParseExcel;

    my $parser = Spreadsheet::ParseExcel->new(
        CellHandler => \&cell_handler,
        NotSetCell  => 1
    );

    my $workbook = $parser->parse("$ARGV[0]");

    sub cell_handler {

        my $workbook    = $_[0];
        my $sheet_index = $_[1];
        my $row         = $_[2];
        my $col         = $_[3];
        my $cell        = $_[4];

        # Do something useful with the formatted cell value
        print $cell->value(), "\n";

    }
