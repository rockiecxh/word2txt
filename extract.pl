#!/usr/bin/perl -w

#require "convertword.pl";
use strict;
use warnings;
use Cwd;
use File::Spec::Functions qw( catfile );


# ת����ʽ
print "ת��WordΪ�ı��ļ�...\n";
#&convertWord();
`perl convertword.pl`;

# �ؼ�������

#my @KEY_WORD = ("���벿��","����ʱ��","������");
my @KEY_WORD = ();

# ��ȡ��ǰĿ¼������·����Ϣ��
my $currentDir = getcwd;
print "��ǰĿ¼: ".$currentDir."\n";

# �����ļ�
my $cfgFile = "$currentDir/�ؼ���.txt";

# ����Դ�ļ���
my $SRC_PATH = "$currentDir/TXT/";

# ������ʱ�ļ���
my $TMP_PATH = "$currentDir/Temp/";

# ����Ŀ���ļ���
my $DEST_PATH="$currentDir/Result/";

# ���Ŀ���ļ��в����ڣ��򴴽��ļ���
mkdir( $DEST_PATH, 0777 ) if ( !-d $DEST_PATH);

opendir TEMP, ${SRC_PATH} or die "�޷���".$SRC_PATH."Ŀ¼������һ��Ŀ¼�Ƿ���ڣ�";

@KEY_WORD=&readConfig();

# ��ȡĿ¼�������ļ�
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

# ���Ŀ���ļ��в����ڣ��򴴽��ļ���
mkdir($DEST_PATH, 0777) if(! -d $DEST_PATH);

# ɾ�����½���ʱĿ¼
deldir($TMP_PATH) if( -d $TMP_PATH);
mkdir($TMP_PATH, 0777) if(! -d $TMP_PATH);

open(FH_PS, ">$destFile") or die "�޷����ļ�: $destFile <$!>";

# ����ļ�ͷ
my $header = (join ',', @KEY_WORD).",";
print FH_PS $header."\n";

foreach (@filelist) {

	next if /^\./; 
	$count=$count+1;
	$srcFile=$SRC_PATH.$_;
	$srcFileBackup=$TMP_PATH.$_.'.bkp';

    print "\n";

	open(FH_B,">$srcFileBackup") or die "�޷����ļ�: $srcFileBackup";
	open(FH,$srcFile) or die "�޷����ļ�: $srcFile";

	while(<FH>){
		#chomp;
		
		s/\r/\|/g;
		s/\n\n/\n/g;
		#s/\n\n/\n/g;
		
		print FH_B $_;
		#s/�ύ��\r/\n�ύ��|/;
		#s/\r�ύ��λ\r/\n�ύ��λ|/;
	
	}
	close(FH_B);
	close(FH);
	
	open(FH_B,$srcFileBackup) or die "�޷����ļ�: $srcFileBackup";

	my $result='';
	while(<FH_B>){
		chomp;
		$line = $_;
		foreach my $key(@KEY_WORD){
			# ���ݹؼ���ƥ����Ӧ��ֵ
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
# ɾ����ʱĿ¼
deldir($TMP_PATH) if( -d $TMP_PATH);
print "All files are processed.\nTotal file processed: $count\n";


sub readConfig(){
	# ��ȡ����ļ� - START
	open(FH_CONFIG, "<$cfgFile") || die "�޷��������ļ� : $cfgFile <$!>";
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
	# ��ȡ����ļ� - END
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