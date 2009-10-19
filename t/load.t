BEGIN {
	@classes = qw(
		Pod::WordML
		Pod::WordML::AddisonWesley
		);
	}

use Test::More tests => scalar @classes;

foreach my $class ( @classes )
	{
	print "bail out! $class did not compile\n" unless use_ok( $class );
	}
