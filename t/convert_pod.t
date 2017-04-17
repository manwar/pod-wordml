#!perl
use strict;
use warnings;

use Test::More;

use File::Spec::Functions;
require './t/lib/transform_file.pl';

chdir 'test-corpus';
my @files = glob( '*.pod' );
chdir '..';

foreach my $file ( @files ) {
	subtest "$file" => sub { transform_file( $file ) };
	}

done_testing();
