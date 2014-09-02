#!/usr/bin/perl -w

#require "convertword.pl";
use strict;
use warnings;
use Cwd;
use File::Spec::Functions qw( catfile );


# 转换格式
print "转换Word为文本文件...\n";
#&convertWord();
`perl convertword.pl`;

# 关键字设置

#my @KEY_WORD = ("申请部门","申请时间","申请人");
my @KEY_WORD = ();

# 获取当前目录的完整路径信息。
my $currentDir = getcwd;
print "当前目录: ".$currentDir."\n";

# 配置文件
my $cfgFile = "$currentDir/关键字.txt";

# 设置源文件夹
my $SRC_PATH = "$currentDir/TXT/";

# 设置临时文件夹
my $TMP_PATH = "$currentDir/Temp/";

# 设置目标文件夹
my $DEST_PATH="$currentDir/Result/";

# 如果目标文件夹不存在，则创建文件夹
mkdir( $DEST_PATH, 0777 ) if ( !-d $DEST_PATH);

opendir TEMP, ${SRC_PATH} or die "无法打开".$SRC_PATH."目录，请检查一下目录是否存在！";

@KEY_WORD=&readConfig();

# 读取目录下所有文件
my @filelist = readdir TEMP; 

my $srcFile='';
my $srcFileBackup='';
my $destFile=$DEST_PATH.'Result.csv';

my $dateflag=0;
my $ownerflag=0;

my $date='';
my $owner='';

my $count=0;
my $flag=0;

my $line='';

# 如果目标文件夹不存在，则创建文件夹
mkdir($DEST_PATH, 0777) if(! -d $DEST_PATH);

# 删除并新建临时目录
deldir($TMP_PATH) if( -d $TMP_PATH);
mkdir($TMP_PATH, 0777) if(! -d $TMP_PATH);

open(FH_PS, ">$destFile") or die "无法打开文件: $destFile <$!>";

# 输出文件头
my $header = (join ',', @KEY_WORD).",";
print FH_PS $header."\n";

foreach (@filelist) {

	next if /^\./; 
	$count=$count+1;
	$srcFile=$SRC_PATH.$_;
	$srcFileBackup=$TMP_PATH.$_.'.bkp';

    print "\n";

	open(FH_B,">$srcFileBackup") or die "无法打开文件: $srcFileBackup";
	open(FH,$srcFile) or die "无法打开文件: $srcFile";

	while(<FH>){
		#chomp;
		
		s/\r/\|/g;
		s/\n\n/\n/g;
		#s/\n\n/\n/g;
		
		print FH_B $_;
		#s/提交人\r/\n提交人|/;
		#s/\r提交单位\r/\n提交单位|/;
	
	}
	close(FH_B);
	close(FH);
	
	open(FH_B,$srcFileBackup) or die "无法打开文件: $srcFileBackup";

	my $result='';
	while(<FH_B>){
		chomp;
		$line = $_;
		foreach my $key(@KEY_WORD){
			# 根据关键字匹配相应的值
			if($line=~/$key\|(.+?)\|(.+)/){
				my $value = $1;
				
				print $count.":".$key."|".$value."\n";
				$result = $result.$value.",";
				
			}
		}
	}

	print FH_PS $result."\n";
	close(FH_B);
	

}

close(FH_PS);
# 删除临时目录
deldir($TMP_PATH) if( -d $TMP_PATH);
print "All files are processed.\nTotal file processed: $count\n";


sub readConfig(){
	# 读取配件文件 - START
	open(FH_CONFIG, "<$cfgFile") || die "无法打开配置文件 : $cfgFile <$!>";
	my @words=();
	while(<FH_CONFIG>) {
		chomp;
		next if /^#/;        # skip comments
		next if /^\s*$/;     # skip empty lines
		
		if (/^\s*(.+),(.+)\s*$/){
			@words = split (/,/,$_);
			last;
		}
		
	}
	print "@words\n";

	close FH_CONFIG;
	trimList(@words);
	# 读取配件文件 - END
}


# Perl trim function to remove whitespace from the start and end of the string
sub trim($)
{
	my $string = shift;
	$string =~ s/^\s+//;
	$string =~ s/\s+$//;
	return $string;
}

# Perl trim function to remove whitespace from the start and end of the string
sub trimList($)
{
	my @myList = ();
	foreach(@_){
		push (@myList, trim($_));
	}
	return @myList;
	
}

sub deldir { 

	my($del_dir)=$_[0]; 
	my(@direct); 
	my(@files); 
	opendir (DIR2,"$del_dir"); 
	my(@allfile)=readdir(DIR2); 
	close (DIR2); 
	foreach (@allfile){ 
	if (-d "$del_dir/$_"){ 
	push(@direct,"$_"); 
	} 
	else { 
	push(@files,"$_"); 
	} 
	} 

	my $files=@files; 
	my $direct=@direct; 
	if ($files ne "0"){ 
	foreach (@files){ 
	unlink("$del_dir/$_"); 
	} 
	} 
	if ($direct ne "0"){ 
	foreach (@direct){ 
	&deldir("$del_dir/$_") if($_ ne "." && $_ ne ".."); 
	} 
	} 

	rmdir ("$del_dir"); 
}