use strict;
use warnings;
package Querylet::Output::Excel::XLS;
use parent qw(Querylet::Output);
# ABSTRACT: output querylet results to an Excel file

use Spreadsheet::WriteExcel;

=head1 SYNOPSIS

 use Querylet;
 use Querylet::Output::Excel::XLS;

 database: dbi:SQLite2:dbname=cpants.db

 query:
   SELECT kwalitee.dist,kwalitee.kwalitee
   FROM   kwalitee
   JOIN   dist ON kwalitee.distid = dist.id
   WHERE  dist.author = 'RJBS'
   ORDER BY kwalitee.dist;

 output format: xls
 output file:   cpants.xls

=head1 DESCRIPTION

This module registers an output handler to produce excel files, using
Spreadsheet::WriteExcel.

=method default_type

The default type for Querylet::Output::Excel::XLS is "xls"

=cut

sub default_type { 'xls' }

=method handler

The output handler uses Spreadsheet::WriteExcel to produce an Excel "xls" file.

=cut

sub handler      { \&_as_xls }
sub _as_xls {
	my ($query) = @_;
	my $results = $query->results;
	my $columns = $query->columns;

	my $xls;
  open(my $fh, ">", \$xls)
		or die "couldn't create temporary filehandle for XLS";
  binmode($fh);

  my $workbook = Spreadsheet::WriteExcel->new($fh)
		or die "couldn't create spreadsheet object";

	my $ws = $workbook->add_worksheet('querylet_results');
	$ws->write('A1', [ map { $query->header($_) } @$columns ]);

	my $range = [ map { [ @$_{@$columns} ] } @$results ];
	$ws->write_col('A2', $range);

	$workbook->close;

	return $xls;
}

1;
