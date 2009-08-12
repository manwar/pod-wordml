use ExtUtils::MakeMaker;

require 5.006;

eval "use Test::Manifest 1.21";

WriteMakefile(
	'NAME'	       => 'Pod::WordML',
	'ABSTRACT'     => "Convert Pod to Microsoft's WordML format",
	'VERSION_FROM' => 'lib/WordML.pm',
	'LICENSE'      => 'perl',
	'AUTHOR'       => 'brian d foy <bdfoy@cpan.org>',
	
	'PREREQ_PM'    => { 
		'Pod::PseudoPod'      => '0',
		'Test::LongString'    => '0',
		'Tie::Cycle'          => '0',
		},

	'PM'           => {
		'lib/WordML.pm'     => '$(INST_LIBDIR)/WordML.pm',
		},

	'MAN3PODS'     => {
		'lib/TaggedText.pm' => '$(INST_MAN3DIR)/Pod::WordML.3',
		},

	clean  => { FILES    => q|Pod-WordML-*| },
	);
