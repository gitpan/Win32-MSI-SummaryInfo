# Before `make install' is performed this script should be runnable with
# `make test'. After `make install' it should work as `perl test.pl'

#########################

use Test::Simple tests => 2;

BEGIN { unshift @INC, "." };

use SummaryInfo;

#########################

ok(1, "use Win32::MSI::SummaryInfo"); 

my $file = "test1.msi";

my $info = Win32::MSI::SummaryInfo::new($file);
ok($info, "SummaryInformation created");

undef $info;
