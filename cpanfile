requires 'Moose';
requires 'Win32::OLE';
requires 'Set::Scalar';
requires 'Chart::Math::Axis';
requires 'Path::Class';
requires 'YAML';
requires 'List::MoreUtils';
requires 'perl', '5.008001';

on test => sub {
    requires 'Test::More', 0.88;
};
