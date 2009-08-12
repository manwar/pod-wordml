# $Id$
package Pod::WordML;
use strict;
use base 'Pod::PseudoPod';

use warnings;
no warnings;

use subs qw();
use vars qw($VERSION);

$VERSION = '0.11';

=head1 NAME

Pod::InDesign::WordML - Turn Pod into Microsoft Word's WordML

=head1 SYNOPSIS

	use Pod::InDesign::WordML;

=head1 DESCRIPTION

***THIS IS ALPHA SOFTWARE. MAJOR PARTS WILL CHANGE***

=head2 The style information

This module takes care of most of the tagged text stuff for you, but you'll
want to insert your own style names. The module gets these by calling 
methods to get the style names. You probably want to create an InDesign
document and export it to tagged text to see what you need.

Override these in a subclass.

=cut

=over 4

=item document_header

This is the start of the document that defines all of the styles. You'll need
to override this. You can take this directly from 

=cut

sub document_header  
	{
	<<'HTML';
<ASCII-MAC>
<Version:5><FeatureSet:InDesign-Roman><ColorTable:=<Black:COLOR:CMYK:Process:0,0,0,1>>
<DefineParaStyle:NormalParagraphStyle=<Nextstyle:NormalParagraphStyle>>
HTML
	}

=item head1_style, head2_style, head3_style, head4_style

The paragraph styles to use with each heading level. By default these are
C<Head1Style>, and so on.

=cut

sub head1_style         { 'Head1Style' }
sub head2_style         { 'Head2Style' }
sub head3_style         { 'Head3Style' }
sub head4_style         { 'Head4Style' }

=item normal_paragraph_style

The paragraph style for normal Pod paragraphs. You don't have to use this
for all normal paragraphs, but you'll have to override and extend more things
to get everything just how you like. You'll need to override C<start_Para> to 
get more variety.

=cut

sub normal_para_style   { 'NormalParagraphStyle' }

=item normal_paragraph_style

Like C<normal_paragraph_style>, but for verbatim sections. To get more fancy
handling, you'll need to override C<start_Verbatim> and C<end_Verbatim>.

=cut

sub code_para_style     { 'CodeParagraphStyle'   }

=item inline_code_style

The character style that goes with C<< CE<lt>> >>.

=cut

sub inline_code_style	{ 'CodeCharacterStyle' }

=item inline_url_style

The character style that goes with C<< UE<lt>E<gt> >>.

=cut

sub inline_url_style    { 'URLCharacterStyle'  }

=item inline_italic_style

The character style that goes with C<< IE<lt>> >>.

=cut

sub inline_italic_style { 'ItalicCharacterStyle' }

=item inline_bold_style

The character style that goes with C<< BE<lt>> >>.

=cut

sub inline_bold_style   { 'BoldCharacterStyle' }

=back

=head2 The Pod::Simple mechanics

Everything else is the same stuff from C<Pod::Simple>.

=cut

sub new { $_[0]->SUPER::new() }

sub emit 
	{
	print {$_[0]->{'output_fh'}} $_[0]->{'scratch'};
	$_[0]->{'scratch'} = '';
	return;
	}

sub get_pad
	{
	# flow elements first
	   if( $_[0]{module_flag}   ) { 'module_text'   }
	elsif( $_[0]{url_flag}      ) { 'url_text'      }
	# then block elements
	# finally the default
	else                          { 'scratch'       }
	}

sub start_Document
	{
	$_[0]->{'scratch'} .= $_[0]->document_header; $_[0]->emit;
	}

sub end_Document    { 1 }	

sub start_head1     { $_[0]{'scratch'}  = '<pstyle:' . $_[0]->head1_style . '>'; }
sub end_head1       { $_[0]{'scratch'} .= "\n"; $_[0]->end_non_code_text }

sub start_head2     { $_[0]{'scratch'}  = '<pstyle:' . $_[0]->head2_style . '>'; }
sub end_head2       { $_[0]{'scratch'} .= "\n"; $_[0]->end_non_code_text }

sub start_head3     { $_[0]{'scratch'}  = '<pstyle:' . $_[0]->head3_style . '>'; }
sub end_head3       { $_[0]{'scratch'} .= "\n"; $_[0]->end_non_code_text }

sub start_head4     { $_[0]{'scratch'}  = '<pstyle:' . $_[0]->head4_style . '>'; }
sub end_head4       { $_[0]{'scratch'} .= "\n"; $_[0]->end_non_code_text }

sub end_non_code_text
	{
	my $self = shift;
	
	$self->make_curly_quotes;
	
	#$self->{'scratch'} .= "\n"; 
	$self->emit
	}
	
sub start_Para      
	{ 
	my $self = shift;
	
	$self->{'scratch'}  = '<pstyle:' . $self->normal_para_style . '>'; 
	
	$self->{'in_para'} = 1; 
	}


sub end_Para        
	{ 
	my $self = shift;
	
	$self->{'scratch'} .= "\n";

	$self->end_non_code_text;

	$self->{'in_para'} = 0;
	}

sub start_figure 	{ }

sub end_figure      { }

sub start_Verbatim { $_[0]{'in_verbatim'} = 1; }

sub end_Verbatim 
	{	
	my @lines = split m/^/m, $_[0]{'scratch'};
	
	my $first = shift @lines;
	my $last  = shift @lines;	
	
	$_[0]{'scratch'} =~ s/\n+\z/\n/;
			
	my $style = $_[0]->code_para_style;
	
	$_[0]{'scratch'} =~ s/^/<pstyle:$style>/gm;

	$_[0]{'scratch'} .= "\n";

	$_[0]->emit();

	$_[0]{'in_verbatim'} = 0;
	}

sub start_B  { $_[0]{'scratch'} .= "<CharStyle:" . $_[0]->inline_bold_style . ">" }
sub end_B    { $_[0]{'scratch'} .= "<CharStyle:>" }

sub start_C  { $_[0]{'scratch'} .= "<CharStyle:" . $_[0]->inline_code_style . ">" }
sub end_C    { $_[0]{'scratch'} .= "<CharStyle:>" }

sub start_E { $_[0]{'in_E'} = 1 }
sub end_E   { $_[0]{'in_E'} = 0 }

sub start_F  { }
sub end_F    { }                                                                                                         

sub start_I  { $_[0]{'scratch'} .= "<CharStyle:" . $_[0]->inline_italic_style . ">" }
sub end_I    { $_[0]{'scratch'} .= "<CharStyle:>" }

sub start_M
	{	
	$_[0]{'module_flag'} = 1;
	$_[0]{'module_text'} = '';
	$_[0]->start_C;
	}

sub end_M
	{
	$_[0]->end_C;
	$_[0]{'module_flag'} = 0;
	}

sub start_N { }
sub end_N   { }

sub start_U { $_[0]->start_I }
sub end_U   { $_[0]->end_I   }

sub handle_text
	{
	my( $self, $text ) = @_;
	
	my $pad = $self->get_pad;
		
	$self->escape_text( \$text );
	
	$self->{$pad} .= $text;		
	}

sub escape_text
	{
	my( $self, $text_ref ) = @_;
		
	# escape escape chars. This is escpaing them for InDesign
	# so don't worry about double escaping for other levels. Don't
	# worry about InDesign in the pod.
	$$text_ref =~ s/\\/\\\\/gx;

	# escape < and >, unless it looks like <0xABCD>, in
	# which case it's a wide character annotated as its
	# hex value.
	$$text_ref =~ s/     < (?! 0x[0-9a-f]{4}   > ) /\\</gx;
	$$text_ref =~ s/(?<! <     0x[0-9a-f]{4} ) >   /\\>/gx;
	
	return 1;
	}

sub make_curly_quotes
	{
	my( $self ) = @_;
	
	my $text = $self->{scratch};
	
	require Tie::Cycle;
	
	tie my $cycle, 'Tie::Cycle', [ qw( <0x201C> <0x201D> ) ];

	1 while $text =~ s/"/$cycle/;
		
	# escape escape chars. This is escpaing them for InDesign
	# so don't worry about double escaping for other levels. Don't
	# worry about InDesign in the pod.
	$text =~ s/'/<0x2019>/g;
	
	$self->{'scratch'} = $text;
	
	return 1;
	}
	
BEGIN {
require Pod::Simple::BlackBox;

package Pod::Simple::BlackBox;

sub _ponder_Verbatim {
	my ($self,$para) = @_;
	DEBUG and print " giving verbatim treatment...\n";

	$para->[1]{'xml:space'} = 'preserve';
	foreach my $line ( @$para[ 2 .. $#$para ] ) 
		{
		$line =~ s/^\t//gm;
		$line =~ s/^(\t+)/" " x ( 4 * length($1) )/e
  		}
  
  # Now the VerbatimFormatted hoodoo...
  if( $self->{'accept_codes'} and
      $self->{'accept_codes'}{'VerbatimFormatted'}
  ) {
    while(@$para > 3 and $para->[-1] !~ m/\S/) { pop @$para }
     # Kill any number of terminal newlines
    $self->_verbatim_format($para);
  } elsif ($self->{'codes_in_verbatim'}) {
    push @$para,
    @{$self->_make_treelet(
      join("\n", splice(@$para, 2)),
      $para->[1]{'start_line'}, $para->[1]{'xml:space'}
    )};
    $para->[-1] =~ s/\n+$//s; # Kill any number of terminal newlines
  } else {
    push @$para, join "\n", splice(@$para, 2) if @$para > 3;
    $para->[-1] =~ s/\n+$//s; # Kill any number of terminal newlines
  }
  return;
}

}

BEGIN {

# override _treat_Es so I can localize e2char
sub _treat_Es 
	{ 
	my $self = shift;

	require Pod::Escapes;	
	local *Pod::Escapes::e2char = *e2char_tagged_text;

	$self->SUPER::_treat_Es( @_ );
	}

sub e2char_tagged_text
	{
	package Pod::Escapes;
	
	my $in = shift;
	return unless defined $in and length $in;
	
	   if( $in =~ m/^(0[0-7]*)$/ )         { $in = oct $in; } 
	elsif( $in =~ m/^0?x([0-9a-fA-F]+)$/ ) { $in = hex $1;  }

	if( $NOT_ASCII ) 
	  	{
		unless( $in =~ m/^\d+$/ ) 
			{
			$in = $Name2character{$in};
			return unless defined $in;
			$in = ord $in; 
	    	}

		return $Code2USASCII{$in}
			|| $Latin1Code_to_fallback{$in}
			|| $FAR_CHAR;
		}
 
 	if( defined $Name2character_number{$in} and $Name2character_number{$in} < 127 )
 		{
 		return chr( $Name2character_number{$in} );
 		}
	elsif( defined $Name2character_number{$in} ) 
		{
		# this need to be fixed width because I want to look for
		# it in a negative lookbehind
		return sprintf '<0x%04x>', $Name2character_number{$in};
		}
	else
		{
		return '???';
		}
  
	}
}

=head1 TO DO

=over 4

=item * beef up entity handling in EE<lt>>. I had to override some stuff from Pod::Escapes

=back

=head1 SEE ALSO

L<Pod::PseudoPod>, L<Pod::Simple>

=head1 SOURCE AVAILABILITY

This source is in Github:

	http://sourceforge.net/projects/brian-d-foy/

If, for some reason, I disappear from the world, one of the other
members of the project can shepherd this module appropriately.

=head1 AUTHOR

brian d foy, C<< <bdfoy@cpan.org> >>

=head1 COPYRIGHT AND LICENSE

Copyright (c) 2009, brian d foy, All Rights Reserved.

You may redistribute this under the same terms as Perl itself.

=cut

1;
