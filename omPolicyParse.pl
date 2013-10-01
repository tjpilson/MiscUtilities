#!/usr/bin/perl

## omPolicyParse.pl
## Tim Pilson
## 8/27/2013
## Parse OM policy files
##
## Input:  Data diretory with OM policy files, *_data
## Output: .log: policy data formated to CSV
##         .error: policies with missing TROUBLETICKET flag

use strict;
use Spreadsheet::WriteExcel;
use Getopt::Long;

my $outputDir   = "."; ## Location of data files
my $outputFile  = "omPolicyParse.csv";  ## Parsed data output
my $outputExcel = "omPolicyParse.xls";  ## Parsed data output
my $outputError = "omPolicyParse.error";  ## Policies w/out troubletickets
my $inputDir    = ".";  ## default current path for inputDir (override with command line option)
my $verbose     = 0;

my $usage = "Usage: $0 [options]\n";
$usage   .= "Options: --verbose, -[vV] Print each policy result to STDOUT.\n";
$usage   .= "         --input,   -[iI] Set input data directory for policy files.\n";
$usage   .= "         --output,  -[oO] Set output data directory for results.\n";
$usage   .= "         --help,    -[hH] Print this help message.\n";

## Set Command Line Options
GetOptions ('input=s'  => \$inputDir,
            'i=s'      => \$inputDir,
            'output=s' => \$outputDir,
            'o=s'      => \$outputDir,
            'verbose'  => \$verbose,
            'v'        => \$verbose,
            'help'     => sub { print $usage; exit; },
            'h'        => sub { print $usage; exit; }
) or die($usage);

## Establish output files
open(OUTPUT,">$outputDir/$outputFile") || die("Can't open file: $outputFile\n");
open(OUTPUTERROR,">$outputDir/$outputError") || die("Can't open file: $outputError\n");

## Create Excel output file
my $workbook  = Spreadsheet::WriteExcel->new("$outputDir/$outputExcel") || die "Can't open Excel output file\n";
my $worksheet = $workbook->add_worksheet();

print "Loading directory list of policy files... ";

## Open the directory where policy files are located
opendir(FILES,$inputDir) || die("Can't open $inputDir\n");
my @fileList = readdir(FILES);

print "done\n\n";

## Set counters for totals
my $totalPolicyCount   = 0;
my $advmonitorCount    = 0;
my $logfileCount       = 0;
my $opcmsgCount        = 0;
my $scheduleCount      = 0;
my $conditionCount     = 0;
my $conditionProcessed = 0;
my $formatError        = 0;
my $noTroubleTicket    = 0;

## Write the output file header
print OUTPUT "OpenView Policy ID,Category,Application,Item,Summary,Priority\n";

## Write the Excel output file header
$worksheet->write(0, 0, 'OpenView Policy ID');
$worksheet->write(0, 1, 'Policy Name');
$worksheet->write(0, 2, 'Category');
$worksheet->write(0, 3, 'Application');
$worksheet->write(0, 4, 'Item');
$worksheet->write(0, 5, 'Summary');
$worksheet->write(0, 6, 'Priority');

## Open each file from the directory
foreach my $file (@fileList) {
  next if ( $file !~ /_data/ );  ## skip non-data files
  (my $id = $file) =~ s/_data//;  ## strip off _data from file name
  $totalPolicyCount++;  ## keep track of how many policies we read

  open(POLICYFILE,"$inputDir/$file") || die("Can't open file: $file\n");

  my $loopstart = 0;  ## Used to determine the start of a condition
  my $innerLoopstart = 0;  ## Used to determine the inner loop structure of a condition
  my $policyType = "";  ## Reset the policy type for new file
  my $policyName = "";  ## Reset the policy name
  my ($application,$msggrp,$object,$text,$severity);

  ## Read each line from the file
  foreach my $polfile (<POLICYFILE>) {
    chomp $polfile;
    $polfile =~ s/^\s+//;  ## Remove initial whitespace

    ## Determine the policy type
    if ( $polfile =~ /^ADVMONITOR/ ) {
      $policyType = "advmonitor";
      $policyName = $polfile; 
      $policyName =~ s/ADVMONITOR\s+//;  ## Remove keyword ADVMONITOR and trailing whitespace
      $policyName =~ s/\"//g;  ## Remove parans
      $advmonitorCount++;
    } elsif ( $polfile =~ /^LOGFILE/ ) {
      $policyType = "logfile";
      $policyName = $polfile; 
      $policyName =~ s/LOGFILE\s+//;  ## Remove keyword LOGFILE and trailing whitespace
      $policyName =~ s/\"//g;  ## Remove parans
      $logfileCount++;
    } elsif ( $polfile =~ /^OPCMSG/ ) {
      $policyType = "opcmsg";
      $policyName = $polfile; 
      $policyName =~ s/OPCMSG\s+//;  ## Remove keyword OPCMSG and trailing whitespace
      $policyName =~ s/\"//g;  ## Remove parans
      $opcmsgCount++;
    } elsif ( $polfile =~ /^SCHEDULE/ ) {
      $policyType = "schedule";
      $policyName = $polfile; 
      $policyName =~ s/SCHEDULE\s+//;  ## Remove keyword SCHEDULE and trailing whitespace
      $policyName =~ s/\"//g;  ## Remove parans
      $scheduleCount++;
    }

    ## Only handle AdvMonitor, Logfile, and OpcMsg types
    ## since these are the only policy types that ticket
    if ( $policyType =~ /advmonitor|logfile|opcmsg/ ) {

      if (( $polfile =~ /CONDITION_ID/ ) && ( $loopstart == 1 )) {  ## Missing TROUBLETICKET flag
        if ( $verbose == 1 ) { print "$id: Contains a condition without a TROUBLETICKET flag\n\n"; }
        print OUTPUTERROR "$id:$policyName: Contains a condition without a TROUBLETICKET flag\n";
        $noTroubleTicket++;
      } elsif ( $polfile =~ /CONDITION_ID/ ) {
        $loopstart = 1;  ## Start the condition id loop

        ## Reset variables for new condition
        $severity    = "";
        $application = "";
        $msggrp      = "";
        $object      = "";
        $text        = "";
      }

      ## Read the severity code
      if (( $loopstart == 1 ) && ( $polfile =~ /SEVERITY/ )) {
        if ( $polfile !~ /Normal/i ) {  ## Ignore condition if categorized as "normal"
          $innerLoopstart = 1;
          $conditionCount++;
          $severity = $polfile;
	  $severity =~ s/SEVERITY\s+//;
        } else {
          $innerLoopstart = 0;  ## kill the loop initializer if a "Normal" condition
        }
      }

      ## Read the application name
      if (( $loopstart == 1 ) && ( $innerLoopstart == 1) && ( $polfile =~ /APPLICATION/ )) {
        $application = $polfile;
	$application =~ s/APPLICATION\s+//;  ## Remove keyword APPLICATION and trailing whitespace
        $application =~ s/\"//g;  ## Remove parans
      }

      ## Read the message group
      if (( $loopstart == 1 ) && ( $innerLoopstart == 1) && ( $polfile =~ /MSGGRP/ )) {
        $msggrp = $polfile;
        $msggrp =~ s/.*(MSGGRP)//;  ## Remove everything before MSGGRP
        $msggrp =~ s/MSGGRP\s+//;  ## Remove whitespace after MSGGRP
        $msggrp =~ s/^\s+//;  ## Remove any remaining initial whitespace
        $msggrp =~ s/\"//g;  ## Remove parans
      }

      ## Read the object
      if (( $loopstart == 1 ) && ( $innerLoopstart == 1) && ( $polfile =~ /OBJECT/ )) {
        $object = $polfile;
	$object =~ s/OBJECT\s+//;  ## Remove keyword OBJECT and trailing whitespace
        $object =~ s/\"//g;  ## Remove parans
      }

      ## Read the message text
      if (( $loopstart == 1 ) && ( $innerLoopstart == 1) && ( $polfile =~ /TEXT/ )) {
        $text = $polfile;
	$text =~ s/TEXT\s+//;  ## Remove keyword TEXT and trailing whitespace
	$text =~ s/^\"//;  ## Remove opening paranthesis
	$text =~ s/\"$//;  ## Remove closing paranthesis
      }

      ## Read the troubleticket line
      if (( $loopstart == 1 ) && ( $innerLoopstart == 1) && ( $polfile =~ /TROUBLETICKET/ )) {
        $loopstart = 0;  ## TROUBLETICKET is the last string so reset loop condition
        $innerLoopstart = 0;  ## TROUBLETICKET is the last string so reset loop condition
        $conditionProcessed++;

        ## Look for missing fields
        if (( $application eq "" ) || ( $msggrp eq "" ) || ( $object eq "" ) || ( $text eq "" )) {
          $formatError++;
        }

	if ( $verbose == 1 ) {
          print "    Policy ID: $id\n";
          print "  Policy Name: $policyName\n";
          print "     Priority: $severity\n";
          print "         Type: $application\n";
          print "     Category: $msggrp\n";
          print "       Object: $object\n";
          print "      Summary: $text\n\n";
        }
        
        print OUTPUT "$id,$policyName,$msggrp,$application,$object,$text,$severity\n";

	## Create data row in Excel output file
        $worksheet->write($conditionProcessed, 0, "$id");
        $worksheet->write($conditionProcessed, 1, "$policyName");
        $worksheet->write($conditionProcessed, 2, "$msggrp");
        $worksheet->write($conditionProcessed, 3, "$application");
        $worksheet->write($conditionProcessed, 4, "$object");
        $worksheet->write($conditionProcessed, 5, "$text");
        $worksheet->write($conditionProcessed, 6, "$severity");
      }
    }
  }
}
close(FILES);
close(OUTPUT);
close(OUTPUTERROR);

if ( $totalPolicyCount == 0 ) {  ## Check if any policies processed
  print "No policy files found\n";
} else {
  ## Total all the policies that were handled
  my $totalHandled = ($advmonitorCount + $logfileCount + $opcmsgCount + $scheduleCount);

  ## Print results/summary
  print "           Total Policies: $totalPolicyCount\n";
  print "         Policies Handled: $totalHandled\n";
  print "        Conditions Parsed: $conditionCount  (found severity codes)\n";
  print "    Conditions (ticketed): $conditionProcessed\n";
  print "Conditions (non-ticketed): $noTroubleTicket\n";
  print "            Format Errors: $formatError  (missing fields)\n\n";

  print "      AdvMonitor Policies: $advmonitorCount\n";
  print "         Logfile Policies: $logfileCount\n";
  print "          OpcMsg Policies: $opcmsgCount\n";
  print "        Schedule Policies: $scheduleCount\n\n";
}


