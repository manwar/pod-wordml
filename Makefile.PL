use ExtUtils::MakeMaker;

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
		'Test::LongString'    => '0',
		'Tie::Cycle'          => '0',
		'Pod::Escapes'        => '0',
		},

	clean  => { FILES    => q|Pod-WordML-*| },
	);
