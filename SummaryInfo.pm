package Win32::MSI::SummaryInfo;

=head1 NAME

Win32::MSI::SummaryInfo - a Perl package to modify MSI Summary Information

=head1 SYNOPSIS

  use Win32::MSI::SummaryInfo;
  
  $info = Win32::MSI::SummaryInfo::new($filename);

  @names = $info->GetNames();

  $subject = $info->GetValue('subject');

  $info->SetValue('pagecount', 100);
  $info->SetValue('subject', 'My subject');
  
  %mtime = (
    'year'        =>  2004,
    'month'       =>  2,
    'wday'        =>  3,
    'day'         =>  3,
    'hour'        =>  15,
    'minute'      =>  35,
    'second'      =>  20,
    'milisecond'  =>  16
  );
  $info->SetValue('createdtm', \%mtime);

  %all = $info->GetAllValues();
  
  $info->Close;
  
=head1 DESCRIPTION

=head2 Summary Information object

It can be obtained using C<MSI::SummaryInfo::new>.
It takes one parameter F<msihandle> or F<filename>. One can obtain Summary Information object either by specifying F<msihandle> to 
an open MSI database or F<filename> of a MSI package.

  $info = Win32::MSI::SummaryInfo::new($filename);

  $info = Win32::MSI::SummaryInfo::new($handle);

This object allows reading and modifying of B<SummaryInformation> stream.

=head1 REMARKS

This module depends on C<Win32::API>, which is used to import the functions out of the F<msi.dll>.

=head1 AUTHOR

Please contact C<empi@post.cz> for questions, bugfixes, and suggestions.

Copyright (c) 2004, Martin Peter. All rights reserved. This program is free 
software; it's published under the BSD License. 
Please see http://www.opensource.org/licenses/bsd-license.php.

=head1 SEE ALSO

Win32::MSI::DB

=cut

require Win32::API;

$VERSION = '1.0';

############ API function definitions
#UINT MsiOpenDatabase(LPCTSTR szDatabasePath, LPCTSTR szPersist, MSIHANDLE* phDatabase);
#my $MsiOpenDatabase = new Win32::API('msi.dll', 'MsiOpenDatabase','PPP','N') || die $!;

#UINT MsiCloseHandle(MSIHANDLE hAny);
my $MsiCloseHandle = new Win32::API('msi.dll', 'MsiCloseHandle', 'N', 'N') || die $!;

# UINT MsiGetSummaryInformation(MSIHANDLE hDatabase, LPCTSTR szDatabasePath, UINT uiUpdateCount, MSIHANDLE* phSummaryInfo);
my $MsiGetSummaryInformation = new Win32::API('msi.dll', 'MsiGetSummaryInformation', 'NPNP', 'N') || die $!;

# UINT MsiSummaryInfoGetPropertyCount(MSIHANDLE hSummaryInfo, UINT* puiPropertyCount);
#$MsiSummaryInfoGetPropertyCount = new Win32::API('msi.dll', 'MsiSummaryInfoGetPropertyCount', 'NP', 'N') || die $!;

# UINT MsiSummaryInfoGetProperty(MSIHANDLE hSummaryInfo, UINT uiProperty, UINT* puiDataType, INT* piValue, FILETIME* pftValue, LPTSTR szValueBuf, DWORD* pcchValueBuf);
my $MsiSummaryInfoGetProperty = new Win32::API('msi.dll', 'MsiSummaryInfoGetProperty', 'NNPPPPP', 'N') || die $!;

# UINT MsiSummaryInfoSetProperty(MSIHANDLE hSummaryInfo, UINT uiProperty, UINT uiDataType, INT iValue, FILETIME* pftValue, LPTSTR szValue);
my $MsiSummaryInfoSetProperty = new Win32::API('msi.dll', 'MsiSummaryInfoSetProperty', 'NNNNPP', 'N') || die $!;

# UINT MsiSummaryInfoPersist(MSIHANDLE hSummaryInfo);
my $MsiSummaryInfoPersist = new Win32::API('msi.dll', 'MsiSummaryInfoPersist', 'N', 'N') || die $!;

#UINT MsiDatabaseCommit(MSIHANDLE hDatabase);
#my $MsiDatabaseCommit = new Win32::API('msi.dll', 'MsiDatabaseCommit', 'N', 'N');

########### API constants definitions
#$MSIDBOPEN_READONLY = 0;
#$MSIDBOPEN_TRANSACT = 1;
#$MSIDBOPEN_DIRECT = 2;
#$MSIDBOPEN_CREATE = 3;
#$MSIDBOPEN_CREATEDIRECT = 4;

my $VT_NULL = 1;
my $VT_I2 = 2;
my $VT_I4 = 3;
my $VT_LPSTR = 30;
my $VT_FILETIME = 64;

########### SummaryInformation stream tags
my %SumInfoTags = (
  'codepage' => [1, $VT_I2], 
  'title' => [2, $VT_LPSTR],
  'subject' => [3, $VT_LPSTR], 
  'author' => [4, $VT_LPSTR],
  'keywords' => [5, $VT_LPSTR], 
  'comments' => [6, $VT_LPSTR],
  'template' => [7, $VT_LPSTR], 
  'lastauthor' => [8, $VT_LPSTR],
  'revnumber' => [9, $VT_LPSTR], 
#  'null' => [10, $VT_NULL],
  'lastprinted' => [11, $VT_FILETIME], 
  'createdtm' => [12, $VT_FILETIME],
  'lastsavedtm' => [13, $VT_FILETIME], 
  'pagecount' => [14, $VT_I4],
  'wordcount' => [15, $VT_I4], 
  'charcount' => [16, $VT_I4],
#  'null' => [17, VT_NULL], 
  'appname' => [18, $VT_LPSTR],
  'security' => [19, $VT_I4]
);

########### Methods
sub new {
  my($hDB_or_filename) = @_;
  my(%info, $hSI);

  $hSI = "\0\0\0\0";

  if (-e $hDB_or_filename) {
    $MsiGetSummaryInformation->Call(0, $hDB_or_filename, 30000, $hSI) && return undef;
  }
  elsif ($hDB_or_filename) {
    $MsiGetSummaryInformation->Call($hDB_or_filename, 0, 30000, $hSI) && return undef;
  }
  else {
    return undef;
  }
  
  $info{'handle'} = unpack('l', $hSI);

  bless(\%info, 'Win32::MSI::SummaryInfo');
}

sub DESTROY {
  my $self = shift;

  $MsiSummaryInfoPersist->Call($self->{'handle'}) && return undef;
  $MsiCloseHandle->Call($self->{'handle'}) && return undef;
  
  $self={};
}

sub Close {
  my $self = shift;
  $self->DESTROY;
}

sub GetNames {
  return keys %SumInfoTags;
}

sub GetValue {
  my ($self, $name) = @_;
  
  my $id = ${$SumInfoTags{$name}}[0];
  my $type = ${$SumInfoTags{$name}}[1];
  
  return undef unless $id;

  my $rtype = "\0\0\0\0";
  my $value = "\0\0\0\0";
  my $filetime = "\0" x 8;
  my $buf = "\0" x 1024;
  my $dbuf = 1024;
  $MsiSummaryInfoGetProperty->Call($self->{'handle'}, $id, $rtype, $value, $filetime, $buf, $dbuf) && return undef;
  
  if (($type == $VT_I2) or ($type == $VT_I4)) {
    return unpack('l', $value);
  }
  elsif ($type == $VT_LPSTR) {
    $buf =~ s/\0+//g;
    return $buf;
  }
  elsif ($type == $VT_FILETIME) {
    my $systemtime = pack('SSSSSSSS', 0, 0, 0, 0, 0, 0, 0, 0);
    my $FileTimeToSystemTime = new Win32::API('kernel32.dll', 'FileTimeToSystemTime', 'PP', 'N') || die $!;
    $FileTimeToSystemTime->Call($filetime, $systemtime);
    my @systemtime =unpack('SSSSSSSS', $systemtime);
    return (
      'year'      =>  $systemtime[0],
      'month'     =>  $systemtime[1],
      'wday'      =>  $systemtime[2],
      'day'       =>  $systemtime[3],
      'hour'      =>  $systemtime[4],
      'minute'    =>  $systemtime[5],
      'second'    =>  $systemtime[6],
      'milisecond'=>  $systemtime[7]
    );


  }
  
  return undef;
}

sub SetValue {
  my ($self, $name, $value) = @_;
  
  my $id = ${$SumInfoTags{$name}}[0];
  my $type = ${$SumInfoTags{$name}}[1];
  
  return undef unless $id;

  if (($type == $VT_I2) or ($type == $VT_I4)) {
    $return = $MsiSummaryInfoSetProperty->Call($self->{'handle'}, $id, $type, $value, 0, 0) && return;
  }
  elsif ($type == $VT_LPSTR) {
    $MsiSummaryInfoSetProperty->Call($self->{'handle'}, $id, $type, 0, 0, $value) && return;
  }
  elsif ($type == $VT_FILETIME) {
    my $systemtime = pack('SSSSSSSS', $value->{'year'}, $value->{'month'}, $value->{'wday'}, 
      $value->{'day'}, $value->{'hour'}, $value->{'minute'}, $value->{'second'}, $value->{'milisecond'});
    my $filetime = pack('LL', 0, 0);
    
    my $SystemTimeToFileTime = new Win32::API('kernel32.dll', 'SystemTimeToFileTime', 'PP', 'N') || die $!;
    $SystemTimeToFileTime->Call($systemtime, $filetime);
    
    $MsiSummaryInfoSetProperty->Call($self->{'handle'}, $id, $type, 0, $filetime, 0) && return;
  }
  return;
}

sub GetAllValues {
  my $self = shift;
  
  my %hash = ();
  foreach ($self->GetNames()) {
    if ($SumInfoTags{$_}[1] != $VT_FILETIME) {
      $hash{$_} = $self->GetValue($_);
    }
    else {
      my %time = $self->GetValue($_);
      $hash{$_} = sprintf('%4d/%02d/%02d  %02d:%02d:%02d', $time{'year'}, $time{'month'}, $time{'day'}, 
        $time{'hour'}, $time{'minute'}, $time{'second'});
    }
  }
  
  return %hash;
}