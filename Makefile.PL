use ExtUtils::MakeMaker 6.64;

require 5.010;

eval "use Test::Manifest 1.21";

WriteMakefile(
	'NAME'	       => 'Pod::WordML',
	'ABSTRACT'     => "Convert Pod to Microsoft's WordML format",
	'VERSION_FROM' => 'lib/Pod/WordML.pm',
	'LICENSE'      => 'perl',
	'AUTHOR'       => 'brian d foy <bdfoy@cpan.org>',
	
	'PREREQ_PM'    => { 
		'Pod::PseudoPod'      => '0',
		'Tie::Cycle'          => '0',
		'Pod::Escapes'        => '0',
		},

	'CONFIGURE_REQUIRES' => {
		'ExtUtils::MakeMaker' => '6.64',
		},

	'TEST_REQUIRES' => {
		'Test::LongString'    => '0',
		},

	'META_MERGE' => {
		'meta-spec' => { version => 2 },
		resources => {
			repository => {
				type => 'git',
				url  => 'git@github.com:briandfoy/Pod-WordML.git',
				web  => 'https://github.com/briandfoy/Pod-WordML',
				},
			},
		no_index => {
			package   => [ qw( Pod::Escapes Pod::Simple::BlackBox ) ],
			},
		},

	clean  => { FILES    => q|Pod-WordML-*| },
	);
