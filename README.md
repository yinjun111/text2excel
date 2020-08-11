text2excel
version: 0.3

Andrew P. Hodges, Ph.D., Jun Yin, Ph.D.
Copyright, Sanford Burnham Prebys Medical Discovery Institute

Description: Perl script to generate compile xlsx file from individual text files. The script will prevent automatic format changes in Excel for gene names, special texts, and combine different text files in to Excel tabs.\n

Usage: 

    #Simple usage
    perl text2excel.pl -i file1.txt,file2.txt -o result.xlsx

    #Set sheet names, and choose different themes
    perl text2excel.pl -i file1.txt,file2.txt -n ShName1,Shname2 -t 1,2 --theme theme2 -o result.xlsx\n

    #Use wildcards for input files
    perl text2excel.pl -i \"*.txt\" --theme theme2 -o result.xlsx\n


Parameters:
	--in|-i           input file(s) separated by \",\", support wildcards
                           if wildcards are used, --names won't be supported.
	--out|-o          output file

	--names|-n        Sheet names
	--txt|-t          column number starting from 0 that should be txt not general separated by \",\"

	--boldfrow|-bfr   bold first row [F]
	--boldfcol|-bfc   bold first column [F]
	--color|-c        color theme to use

    #themes to be used
	--theme           theme1, by AH 
                      theme2, by JY, adding wrap text, filter etc. Now default option. [theme2]
                      theme0, don't change format

	--delim|-d        default is tab-delimited; use '' for other entries [\\t]
	--verbose|-v      Verbose\n	
