use Test::More tests => 2;

BEGIN {
  use_ok('Querylet::Query');
  use_ok('Querylet::Output::Excel::XLS');
}

diag( "Testing  $Querylet::Output::Excel::XLS::VERSION" );
