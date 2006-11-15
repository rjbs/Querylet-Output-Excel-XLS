package Querylet::Output::Excel::XLS;
use base qw(Querylet::Output);

use warnings;
use strict;

=head1 NAME

Querylet::Output::Excel::XLS - output querylet results to an Excel file

=head1 VERSION

version 0.12

 $Id$

=cut

our $VERSION = '0.13';

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

=over 4

=item C<< default_type >>

The default type for Querylet::Output::Excel::XLS is "xls"

=cut

sub default_type { 'xls' }

=item C<< handler >>

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

=back

=head1 AUTHOR

Ricardo SIGNES, C<< <rjbs@cpan.org> >>

=head1 BUGS

Please report any bugs or feature requests to
C<bug-querylet-output-text@rt.cpan.org>, or through the web interface at
L<http://rt.cpan.org>.  I will be notified, and then you'll automatically be
notified of progress on your bug as I make changes.

=head1 COPYRIGHT

Copyright 2004 Ricardo SIGNES, All Rights Reserved.

This program is free software; you can redistribute it and/or modify it
under the same terms as Perl itself.

=cut

1;
