use ExtUtils::MakeMaker;
WriteMakefile(
    'NAME'	   => 'Win32::MSI::SummaryInfo',
    'VERSION_FROM' => 'SummaryInfo.pm',
    'PREREQ_PM'	   => {
        'Win32::API' => 0,
      },
    'NORECURS'     => 1,
  ($] ge '5.005') ? (
    'AUTHOR'       => 'Martin Peter (empi@post.cz)',
    'ABSTRACT'     => 'Interface for MSI SummaryInformation stream',
  ) : (),
);
