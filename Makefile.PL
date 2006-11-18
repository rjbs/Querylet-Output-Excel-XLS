use strict;
use warnings;
use ExtUtils::MakeMaker;

WriteMakefile(
    NAME                => 'Querylet::Output::Excel::XLS',
    AUTHOR              => 'Ricardo SIGNES <rjbs@cpan.org>',
    VERSION_FROM        => 'lib/Querylet/Output/Excel/XLS.pm',
    ABSTRACT_FROM       => 'lib/Querylet/Output/Excel/XLS.pm',
    LICENSE             => 'perl',
    PL_FILES            => {},
    PREREQ_PM => {
        'Test::More'  => 0,
        'Querylet::Query'  => 0.28,
        'Querylet::Output' => 0.24,
        'Spreadsheet::WriteExcel' => 2.04,
    },
    dist                => { COMPRESS => 'gzip -9f', SUFFIX => 'gz', },
    clean               => { FILES => 'Querylet-Output-Excel-XLS-*' },
);
