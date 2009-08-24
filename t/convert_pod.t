#!perl
use strict;
use warnings;

use Test::More 'no_plan';

my @files = glob( catfile( qw(test-corpus *.pod) ) );

foreach my $file ( @files )
	{
	my $parser = Pod::WordML->new;
	
	my $string;
	open my($fh), '>', \ $string;
	
	$parser->output_fh( $fh );
	
	$parser->parse_file( $file );
	
	diag( $string );
	}
