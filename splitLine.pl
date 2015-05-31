# 引数に分割する行数を指定 

use strict;
use warnings;

my $file = "test.txt"; 
my $MaxLine = $ARGV[0];
my $LineCount = 0;
my $FileCount = 0;

open(my $fhRead, "<", $file)
  or die "Cannot open $file: $!";

while(my $line = readline $fhRead){ 
	if($LineCount >= $MaxLine || $LineCount == 0) {
		$LineCount = 1;
		$FileCount++;
		open(OUT, ">file" . $FileCount . ".txt");
	}
	else {
		$LineCount++;
	}
	
	chomp $line;

	print OUT $line, "\n";
}

close($fhRead);
close(OUT);
