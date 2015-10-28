#===============================================================================
# This PERL script is designed to extract the following information from the 
# LDRA TCF file (Low Level Test file)(*.tcf)(Module testing file) and Result 
# file (.htm)used in the software testing domain.
#
# From TCF (*.tcf) file:
# ----------------------
#  1) Test Script File name
#  2) Test Script File version (latest version)
#  3) Requirement of the module
#  4) Source file of the module
#  5) Module/Function name (Unit Under Test)
#  6) Total number of the Testcases
#  7) Software Data Dictionary Baseline information
#  8) LDRA Testbed Version
#  9) Software Build Version
#  10) Station ID (The hostname of the PC on which the script was executed)
#
# From Result (*_Result.htm) file:
# ----------------------------------
#  1) Test case PASS/FAIL/SUSPENDED information 
#  2) Overall Status  
#
# From Coverage (*_Coverage.htm) file:
# ------------------------------------
#  1) Statement coverage of function under test
#  2) Branch coverage of function under test 
#
# Other Information:
# ------------------
#  1) Presence of Manual inspection
#  2) Presence of Problem Report (PR) document
#
# Output: The above information will be presented through an EXCEL file
#
#===========================================================================================

#===========================================================================================
# Revision History
#===========================================================================================
#   Rev     Author            Created/Modified   Remarks           
#                                  Date             
#===========================================================================================
#   1.0     Archana S          2-Aug-2015       Updated to extrac the TCF's latest version
#                       	                    information 
#
#===========================================================================================

#!/usr/bin/perl -w

BEGIN{push @INC, 'C:\\Perl\\lib';}
use Text::Tabs;
use Cwd;
use Win32;
use Win32::OLE;
use Win32::OLE qw(in with);
use Win32::OLE::Variant;
use Win32::OLE::Const 'Microsoft Excel';
use POSIX qw(strftime);
use Data::Dumper;
use HTML::TableExtract;
use Text::Table;

Win32::MsgBox ("Extracting the LLT Test artifacts data",
               0,
			   TCF_Report);

my $path = getcwd();	
$path =~ s/\//\\/g;

print "Enter Tcf folder path: ";	   
$Current_Directory = <>;
chomp($Current_Directory);
$Current_Directory =~ s/\//\\/g;

# Creating Excel
my $Excel = Win32::OLE->new('Excel.Application', 'Quit');
   $Excel->{'Visible'} = 1;

# Get already active Excel application or open new 
my $Excel = Win32::OLE->GetActiveObject('Excel.Application') 
         || Win32::OLE->new('Excel.Application', 'Quit');   
   $Excel->{'Visible'} = 1;        #0 is hidden, 1 is visible
   $Excel->{DisplayAlerts} = 0;    #0 is hide alerts

# Create Excel file with 1 worksheets
$Excel->{SheetsInNewWorkBook} = 1;
$TestReport = $Excel->Workbooks->Add();

#my $trans_format = $Excel->add_format();
#$trans_format->set_num_format('C:C','$0.0');

# Naming the first sheet as "TCF Report"
$Sheet_1 = $TestReport->Worksheets(1);
$Sheet_1->{Name} = "TCF Report";

#$trans_format->set_num_format('$0.00');
#$Sheet_1->set_column('C:C', undef, $trans_format);

# access Word application
my $msword = Win32::OLE->new("Word.Application",'Quit')|| die "Couldn't run Word";

#Writing the Sheet header (first row)
$Row_Index = 1;
$Sheet_1->Range("A$Row_Index:V$Row_Index")->Interior->{ColorIndex} = 6;
$Sheet_1->Cells($Row_Index, 1)->{Value} = "Sl. No.";
$Sheet_1->Cells($Row_Index, 2)->{Value} = "TCF File names";
$Sheet_1->Cells($Row_Index, 3)->{Value} = "TCF Version";
$Sheet_1->Cells($Row_Index, 4)->{Value} = "Requirements";
$Sheet_1->Cells($Row_Index, 5)->{Value} = "Source file";
$Sheet_1->Cells($Row_Index, 6)->{Value} = "Functions";
$Sheet_1->Cells($Row_Index, 7)->{Value} = "Testcase Count";
$Sheet_1->Cells($Row_Index, 8)->{Value} = "SDD Baseline";
$Sheet_1->Cells($Row_Index, 9)->{Value} = "LDRA Testbed Version";
$Sheet_1->Cells($Row_Index, 10)->{Value} = "Build Version";
$Sheet_1->Cells($Row_Index, 11)->{Value} = "Station ID";
$Sheet_1->Cells($Row_Index, 12)->{Value} = "No. of Test Cases PASSED";
$Sheet_1->Cells($Row_Index, 13)->{Value} = "No. of Test Cases FAILED";
$Sheet_1->Cells($Row_Index, 14)->{Value} = "No. of Test Cases SUSPENDED";
$Sheet_1->Cells($Row_Index, 15)->{Value} = "Overall status";
$Sheet_1->Cells($Row_Index, 16)->{Value} = "Any Manual Inspection?";
$Sheet_1->Cells($Row_Index, 17)->{Value} = "Statement Coverage";
$Sheet_1->Cells($Row_Index, 18)->{Value} = "Branch Coverage";
$Sheet_1->Cells($Row_Index, 19)->{Value} = "Test cases which are passed";
$Sheet_1->Cells($Row_Index, 20)->{Value} = "Test cases which are failed";
$Sheet_1->Cells($Row_Index, 21)->{Value} = "PR present?";
$Sheet_1->Cells($Row_Index, 22)->{Value} = "Suspended test cases";
			   
my $fileCnt = 0;

# Variables to hold the requirement information from the TCF files
my @req_text1_tcf; # Holds the splitted requirement text from the TCF file name
my @req_text2_tcf; # Holds the splitted requirement text along with TCF file's extension
my $req_tag_tcf = 0; # Holds the actual requirement number
my $req_str_tcf = 0;

# Variables to hold the requirement information from the Result files
my @req_text1_result; # Holds the splitted requirement text from the Result file name
my $req_tag_result = ""; # Holds the actual requirement number
my $req_str_result = "";
my %TC_Status;

# Variables to hold the requirement information from the Coverage files
my $headers = [ 'Procedure', 'Statement', 'Branch' ];
my $table_extract = HTML::TableExtract->new(headers => $headers);
my @tables_array = "";
my $function_name = "";
my $req_TCF_tag_for_Coverage = "";

#open the directory where the files are placed
opendir DIR, $Current_Directory or die "cannot open dir $Current_Directory: $!";
@file_of_dir= readdir DIR;
closedir DIR;

# Variables related to extraction of the TCF version info
my @RevHisdata;
my $RevHisCnt = 0;
my $TCFVerNum = 0;

#iterate through each file   
foreach $file (@file_of_dir)
{

	$Sub_path="$Current_Directory\\$file";

	next if ($file =~ m/^\./);
	%TC_Status=();
 
	if($Sub_path !~ m/\.pl$/ig)
	{   
		# enter into the file to read all artifacts
		opendir D, $Sub_path or die "cannot open dir $Sub_path: $!";
		@file_of_dir_sub= readdir D;
		closedir D;
	 
		$tcPassCounter = 0;
	    $tcFailCounter = 0;
	    $only_MI_Flag=0;
		
		# iterate through each artifact of the current file
		foreach $artifact (@file_of_dir_sub)
	    {
		    next if ($artifact =~ m/^\./);
			
			# enter into the test case file(.tcf)
			if ($artifact =~ m/.*.tcf$/)
			{
				$only_MI_Flag=1;
				# Extract the count of test cases
				tcf_file($artifact);
			}
			
			# enter into the Result file(Result.htm)
			elsif ($artifact =~ m/.*_Results.*\.htm$/)
			{
	            # Extract the Test case's PASS/FAIL information
				tcf_result_info($artifact);
			}

			# enter into the Coverage file(Coverage.htm)			
			elsif ($artifact =~ m/.*Coverage.htm$/)
			{
				# Extract the Coverage (Statement and Branch) information of UUT
				tcf_coverage_info($artifact, $function_name);
			}
			
			# Check if manual inspection is present for the current UUT
			elsif ($artifact =~ m/.*_Manual_Inspection.doc/)
			{
			    #write "YES" if MI is present
				$Sheet_1->Cells($Row_Index, 16)->{Value} = "Yes";
				
			    # Check if Test is performed only through MI
				if($only_MI_Flag == 0)
				{
				    # Extract the Manual Inspection information of UUT
					manual_inpection_info();
				}
			}
			
            # Check if Problem Report is raised for the current UUT
			elsif ($artifact =~ m/.*_PR.doc/)
			{
				$Sheet_1->Cells($Row_Index, 21)->{Value} = "Yes";
			}
        }
		
		#read the status of test cases 
		for ($index=1;$index<=$tc_count;$index++)
		{
			#Check if any of the test case is suspended and never executed
			if($TC_Status{$index}=~/SUSPENDED/)
			{
				#Write the suspended test case number
				$Sheet_1->Cells($Row_Index, 22)->{Value} .= "$index,";
			}
		}
	}	
}

#-------------------------------------------------------------------------------
# Formatting the Excel spreadsheet
#
#-------------------------------------------------------------------------------
$Sheet_1->Range("A$Row_Index:V$Row_Index") -> Borders(xlEdgeBottom)-> {LineStyle}  = xlContinuous;
$Sheet_1->Range("A$Row_Index:V$Row_Index") -> Borders(xlEdgeBottom)-> {Weight}     = xlThin;
$Sheet_1->Range("C1:V$Row_Index") -> Borders(xlEdgeRight)-> {LineStyle}  = xlContinuous;
$Sheet_1->Range("C1:V$Row_Index") -> Borders(xlEdgeRight)-> {Weight}     = xlThin;
$Sheet_1->Range("A1:V$Row_Index") -> Borders(xlInsideVertical)-> {LineStyle}  = xlContinuous;
$Sheet_1->Range("A1:V$Row_Index") -> Borders(xlInsideVertical)-> {Weight}     = xlThin;
$Sheet_1->Range("A1:V$Row_Index") -> Borders(xlInsideHorizontal)-> {LineStyle}  = xlContinuous;
$Sheet_1->Range("A1:V$Row_Index") -> Borders(xlInsideHorizontal)-> {Weight}     = xlThin;
$Sheet_1->Range("A:V") -> {Columns} -> Autofit;
$Sheet_1->Range("A:V") -> {HorizontalAlignment} = xlHAlignCenter;
$Sheet_1->Range("A:V") -> Font -> {Size} = 10;
$Sheet_1->Range("A:V") -> Font -> {Bold} = 1;
$Sheet_1->Range("A:V") -> Font -> {Name} = "Calibri";
#-------------------------------------------------------------------------------
# Creating and Saving the Excel file with DATE
#
#-------------------------------------------------------------------------------
$Date = POSIX::strftime( "%d_%m_%Y", localtime() );
$TestReport->SaveAs({Filename => "$path\\TCF_Report_$Date.xls", 
					FileFormat => xlWorkbookNormal}); 
$TestReport->close();
$TestReport->Quit();

#-------------------------------------------------------------------------------
# Main loop to fetch the details from the TCF (.tcf) and Result (.htm) files
#
#-------------------------------------------------------------------------------
sub tcf_file
{	
	my ($TCF_File) = $_[0];
	$Requirement = "";
	$SDDBaseline = "";
	$SourceFile = "";
	$Function = "";
	$Testcase = "TEST M";
	$LDRATestbedVer = "";
	$BuildVersion = "";
	$Hostname = "";
	$tc_count = 0;
	$req_tag_tcf = 0;
	$req_str_tcf = 0;
	$function_name = "";
	
	
	$tcf_path = "$Sub_path\\$TCF_File";

	# Increment the file counter.
	$fileCnt++;
	
	# Extract the requirement tag/number from the TCF file name
	@req_text1_tcf = split(/_/, $TCF_File);
	$req_str_tcf = $req_text1_tcf[3];
	
	for( $chop_index = 0 ; $chop_index < length(".tcf"); $chop_index++)
	{
		chop($req_str_tcf);
	}
	
	$req_tag_tcf = $req_str_tcf;
	$req_TCF_tag_for_Coverage = $req_tag_tcf;
	
    open (TCF_FILE, "<$tcf_path") or  
			die "Could not read file $tcf_path : $! \n";			
		
	# Increment the Row index
	$Row_Index++;
		
	# Write the Sl. No. 
	$Sheet_1->Cells($Row_Index, 1)->{Value} = "$fileCnt";		
		
	# Write the TCF file name
	$Sheet_1->Cells($Row_Index, 2)->{Value} = "$TCF_File";			
	
	while (my $line = readline(TCF_FILE))
    {
		# Extract the requirement information
		tcf_file_parser($line, "Requirements:", 4);
			
		# Extract the Source file name
		tcf_file_parser($line, "RelativeFile =", 5);
			
		# Extract the Function name (Unit under test)
		tcf_file_parser($line, "Function:", 6);
			
		# Count the number of the test cases in TCF
		if(($line =~ /$Testcase/))
		{
			$tc_count++;

			# Write the Test case count
			$Sheet_1->Cells($Row_Index, 7)->{Value} = "$tc_count";
		}			
			
		# Extract the SDD baseline information
		tcf_file_parser($line, "Baseline Version:", 8);

		# Extract the LDRA Testbed version information
		tcf_file_parser($line, "GENERATED_BY =", 9);
			
		# Extract the Source build version information
		tcf_file_parser($line, "UNCFile = ", 10);
			
		# Extract the HOSTNAME information 
		#(This will be treated as 'Station ID')
		tcf_file_parser($line, "HOSTNAME =", 11);
	}
	
	# Close each test script file after reading
 	close TCF_FILE;
	
	# ------------------------------------------------------------------------
	# Code block to extract the TCF version information
	
	open (TCF_FILE_HDL, "<$tcf_path") or  
			die "Could not read file $tcf_path : $! \n";
	
	# Loop to extract the Revision History table information
	while(<TCF_FILE_HDL>)
	{
		if(/VERSION INFO/../Description:/)
		{					
			# Push the Revision history information lines into an array
			push(@RevHisdata, $_);
    
			# Loop-through the array till the heading is 'Description:'
			do
			{
				# increment the line number
				$RevHisCnt++;
				
				# check the revision number
				if(/^([0-9].[0-9]).*/)
				{
					# store the latest version number 
					$TCFVerNum = $1;
				}
			}
			while($RevHisdata[$RevHisCnt] =~ m/Description:/);
		}
	}
    
	# Print the TCF version information
	$Sheet_1->Cells($Row_Index, 3)->{Value} = "$TCFVerNum";	
	
	# Close each test script file after reading
 	close TCF_FILE_HDL;
}

#-------------------------------------------------------------------------------
# Function: tcf_file_parser
# 
# Description: Function to parse the TCF file and search the text supplied as
#              as argument, and print the extracted information into EXCEL file
#
# Arguments: parsed line
#            text to search
#            column number (to print)
#-------------------------------------------------------------------------------
sub tcf_file_parser {
	my ($parsed_line, $text_to_search, $colNum) = @_;
	
	$text_to_print = "";
	$value = 0;	
	
	# Extract the requirement information
	if(($parsed_line =~ s/$text_to_search//))
	{
		chomp($parsed_line);
		$text_to_print = $parsed_line;	
		
		# Remove the leading whitespaces
		$text_to_print =~ s/^\s+//;
		# Remove the trailing whitespaces
		$text_to_print =~ s/\s+$//;
		
		# Remove the ".\" from the Source file name
		if(($colNum == 5))
		{
			remove_extra_char($colNum);
		}
		
		# Remove the "()" from the function name
		if(($colNum == 6))
		{
			remove_extra_char($colNum);
		}
		
		# extract SDD build information
		if(($colNum == 10))
		{
	    @Build_info = split(/\\/, $text_to_print);	
		$text_to_print=$Build_info[4];
		}
	
		# Write the Requirement number
		$Sheet_1->Cells($Row_Index, $colNum)->{Value} = "$text_to_print";
	}	
}


#-------------------------------------------------------------------------------
# Function: tcf_result_info
# 
# Parameters: Requirement tag (TCF files' requirement tag)
#
# Description: Function to parse the TCF's Result.htm file and fetch the 
#              PASS, FAIL and SUSPENDED details of the test cases.
#
# Arguments: None
#-------------------------------------------------------------------------------
sub tcf_result_info
{
	my ($req_tag) = $_[0];
	$supended_TC=0;
	$result_flag=0;
	$RESULT_File= "$Sub_path\\$req_tag";
	

	$result_line = "";
	$requirement_no = 0;
			
	$result_flag=1;
	open (RESULTFILE, '<', $RESULT_File) or  
			die "Could not read file $RESULT_File : $! \n";

	while (my $line = readline(RESULTFILE))
	{
		# Extract the requirement information
		if(($line =~ m/HREF = "#tc_link_(\d+)"/))
		{ 
			#get the test case number
			$tcnum=$1;
			chomp($line);
			$result_line = $line;		
			
			#Extract the details of 'passed' test case 
			if(($result_line =~ m/PASS/))
			{
				# increment the test case PASS counter
				$tcPassCounter++;
				
				# List of test case numbers which are 'PASS'
				$Sheet_1->Cells($Row_Index, 19)->{Value} .= "$tcnum,";
				
				# Write the count of test case PASS number
				$Sheet_1->Cells($Row_Index, 12)->{Value} = "$tcPassCounter";
				
				#Set the status of the test case as 'PASS'
				$TC_Status{$tcnum}="PASS";
			}
			
			#Extract the details of 'failed' test case 			
			elsif (($result_line =~ m/FAIL/))
			{
				# increment the test case FAIL counter
				$tcFailCounter++;
				
				# List of test case numbers which are 'FAIL'
				$Sheet_1->Cells($Row_Index, 20)->{Value} .= "$tcnum,";
				
				# Write the count of test case FAIL number
				$Sheet_1->Cells($Row_Index, 13)->{Value} = "$tcFailCounter";
			
				#Set the status of the test case as 'FAIL'			
				$TC_Status{$tcnum}="FAIL";
			}
			
			#Extract the details of 'suspended' test case 			
			elsif (($result_line =~ m/SUSPENDED/))
			{
		
				# check if the test case is neither PASS nor FAIL
				if (($TC_Status{$tcnum} ne "PASS")&&($TC_Status{$tcnum} ne "FAIL"))
				{
					#Set the status of the test case as 'SUSPENDED'						   
					$TC_Status{$tcnum}="SUSPENDED";
				}
			}

			# if there are no PASS test cases, fill '0' (zero)
			if(($tcPassCounter == 0))
			{
				# if there are no passes, fill '0' (zero)
				$Sheet_1->Cells($Row_Index, 12)->{Value} = "0";
			}
			
			# if there are no FAIL test cases, fill '0' (zero)
			if(($tcFailCounter == 0))
			{
				# if there are no failures, fill '0' (zero)
				$Sheet_1->Cells($Row_Index, 13)->{Value} = "0";
			}								
		}
	}
				
	# Close each result file after reading
	close RESULTFILE;
	
	# Get the count of test case 'suspended' number
	$supended_TC=$tc_count-($tcPassCounter+$tcFailCounter);
	
	#Check if suspended test cases are found
	if($supended_TC>0)
	{
		# Write the count of test case 'suspended' number
		$Sheet_1->Cells($Row_Index, 14)->{Value} = "$supended_TC are suspended";
		
		# Write the Overall status of UUT as 'FAIL'		
		$Sheet_1->Cells($Row_Index, 15)->{Value} = "FAIL";
		$Sheet_1->Cells($Row_Index, 15) -> Interior -> {ColorIndex} = 3;
	}
	
	#check if any failed test case is found and update the Overall status of UUT as 'FAIL'
	elsif($tcFailCounter>0)
	{
		$Sheet_1->Cells($Row_Index, 15)->{Value} = "FAIL";
		$Sheet_1->Cells($Row_Index, 14)->{Value} = "NONE";
		$Sheet_1->Cells($Row_Index, 15)->Interior->{ColorIndex} = 3;
	}
	
	#Otherwise update the Overall status of UUT as 'PASS'
	else
	{
		$Sheet_1->Cells($Row_Index, 15)->{Value} = "PASS";
		$Sheet_1->Cells($Row_Index, 14)->{Value} = "NONE";
		$Sheet_1->Cells($Row_Index, 15)->Interior->{ColorIndex} = 4;
	}
}

#-------------------------------------------------------------------------------
# Function: tcf_coverage_info
# 
# Parameters: Requirement tag (TCF files' requirement tag)
#             Function name (Function under test)
#
# Description: Function to parse the TCF's Coverage.htm file and fetch the 
#              'Statement' and 'Branch' coverage information of function under 
#			   test depcited under "Summary of Results for Selected Procedures" 
#              table of <xxx_Coverage.htm> file.
#
# Arguments: None
#-------------------------------------------------------------------------------
sub tcf_coverage_info
{
	my ($req_tag, $func_name) = @_;
	
	#trim the requirement tag
    $req_tag =~ s/^\s+//;
    $req_tag =~ s/\s+$//;
	
	#trim the function name
    $func_name=~ s/^\s+//;
    $func_name=~ s/\s+$//;

	$result_line = "";
	$requirement_no = 0;
	
	# Extract the requirement tag/number from the COVERAGE file name
	@req_text1_result = split(/_/, $req_tag);
	
	# 3rd index will hold the requirement number
	$req_tag_coverage = $req_text1_result[3];	 
	
	TABLE_LOOP:
	{
		$table_extract->parse_file("$Sub_path\\$req_tag");
		@tables_array = $table_extract->tables;

		for my $table (@tables_array) 
		{
			my @rows_array = $table->rows;
			for my $table_row (@rows_array) 
			{
				for my $table_col (@$table_row) 
				{ 
					$table_col =~ s/^\s+//;
					$table_col =~ s/\s+$//;

					if(($table_col =~/\b$func_name\b/))
					{					
						# Write the Statement coverage information
						$Sheet_1->Cells($Row_Index, 17)->{Value} = "@$table_row[1]";

						# Write the Branch coverage information
						$Sheet_1->Cells($Row_Index, 18)->{Value} = "@$table_row[2]";

						# Once the information is found, exit from
						# this table loop.
						last TABLE_LOOP;
					}
				}
			}	
		}
	}			
}
	
#-------------------------------------------------------------------------------
# Function: remove_extra_char
# 
# Description: Function to remove the extra characters from function name '()' 
#			    and Source file name '.\'
#
# Arguments: column number 
#-------------------------------------------------------------------------------
sub remove_extra_char 
{
	my ($colNum) = @_;
	
	# Remove the ".\" from the Source file name
	if(($colNum == 5))
	{
		$text_to_print = substr($text_to_print, 2);
	}		
	
	# Remove the "()" from the function name
	if(($colNum == 6))
	{
	    $text_to_print =~ s/\(//;
		$text_to_print =~ s/\)//;
		
		# Copy the function name
		$function_name = $text_to_print;		
	}
}

#-------------------------------------------------------------------------------
# Function: manual_inpection_info
# 
# Description: Function to extract function name, Requirement tag, Source code 
#              file name and test case count from Manual Inspection
#
# Arguments: None 
#-------------------------------------------------------------------------------
sub manual_inpection_info
{
	my $SC_flag=1;
	my $Req_flag=1;
	my $functionName_flag=1;
	my $sourceFile;
	my $tc_count_in_MI=0;
	
	$Row_Index++;
	$fileCnt++;
	
	#write "YES" if MI is present
	$Sheet_1->Cells($Row_Index, 16)->{Value} = "Yes";

	# Write the Sl. No. 
	$Sheet_1->Cells($Row_Index, 1)->{Value} = "$fileCnt";		

	# Write the TCF file name
	$Sheet_1->Cells($Row_Index, 2)->{Value} = "$artifact";	
	$Sheet_1->Cells($Row_Index, 12)->{Value}= "Test is performed by Manual Inspection";
	
	my $document = $msword->Documents->Open({FileName =>"$Sub_path\\$artifact", ReadOnly => 1})or die $!;
	
	#extracting Manual Inspection data
	for my $paragraph (in $document->Paragraphs)
	{
		
		# Remove trailing ^M (the paragraph marker) from Range.
		my($text) = substr($paragraph->Range->Text, 0, -1);
		
		chomp($text);
		
	   #get the source code name from MI
		if ($SC_flag == 0)
		{  
			if($text=~ /(\S+).c.*/)
			{	
				$sourceFile=$1.".c";
				$Sheet_1->Cells($Row_Index, 5)->{Value} = "$sourceFile";	 
			}
			$SC_flag=1;
		}
		if ($text =~ /Source Code/)
		{
			$SC_flag=0;
		}
		
		#get the Requirement tag from MI
		if ($text =~ /^(\S+)-(\S+)-(LLR|LDR)#(\d+)/)
		{
			if ($Req_flag == 1)
			{
				$Sheet_1->Cells($Row_Index, 4)->{Value} = $text ;
				$Req_flag = 0;
			}
		}
		
		#get the Function Name of UUT from MI
		if ($functionName_flag == 1)
		{
			if ($text=~ /^Function:/)
			{
			
				$text=~s/Function://;
				$text=~s/\(\)$//;
				$functionName_flag=0;
				$Sheet_1->Cells($Row_Index, 6)->{Value} = $text ;
			}
		}
		
		#get the test case count of UUT from MI
		if($text=~/^Test Case M(\d+):/)
		{
			$tc_count_in_MI++;
		}
		
		#get the Baseline Info from MI
		if($text=~/^(\S+) (\S+) SDD(.*)Baseline (.*)/)
		{
			my $code= $2; #variable to get whether the code is BBP or MEP
			my $num= $4;  #variable to get Baseline version of SDD
			
			# Remove the leading whitespaces
			$num =~ s/^\s+//;
			# Remove the trailing whitespaces
			$num =~ s/\s+$//;
			$Sheet_1->Cells($Row_Index, 8)->{Value}= "SDD $code $num";
		}
	}
	
	#write the test case count 
	$Sheet_1->Cells($Row_Index, 7)->{Value} = $tc_count_in_MI ;
	$document->Close();
	$document->Quit();
}

