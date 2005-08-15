#!/usr/bin/perl -w
use Email::Simple;
use Email::Simple::Headers

$file = shift(@ARGV);

open(INFILE, $file) || die "Failed to open $file";

$fileText = "";

while (<INFILE>) {
   $fileText .= $_;
}

close(INFILE);

$email = Email::Simple->new($fileText);

print "Headers:\n";
foreach ($email->headers) {
   print "\t$_: " . $email->header($_) . "\n";
}
print "\n\nBody:\n";
print $email->body. "\n";

