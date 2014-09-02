#!/usr/bin/perl -w
# usage: perl conword.pl srcFolder destFolder

use strict;
use warnings;
use Cwd;
use File::Spec::Functions qw( catfile );

use Win32::OLE;
use Win32::OLE::Const 'Microsoft Word';

$Win32::OLE::Warn = 3;


# ��ȡ��ǰĿ¼������·����Ϣ��
my $currentDir = getcwd;
print "��ǰĿ¼: ".$currentDir."\n";

# ����Դ�ļ���
my $SRC_PATH = "$currentDir/Word/";

# ����Ŀ���ļ���
my $DEST_PATH="$currentDir/TXT/";

deldir($DEST_PATH) if( -d $DEST_PATH);

# ���Ŀ���ļ��в����ڣ��򴴽��ļ���
mkdir( $DEST_PATH, 0777 ) if ( !-d $DEST_PATH);

opendir TEMP, ${SRC_PATH} or die "�޷���".$SRC_PATH."Ŀ¼������һ��Ŀ¼�Ƿ���ڣ�";

# ��ȡĿ¼�������ļ�
my @filelist = readdir TEMP; 

my $srcFile='';
my $destFile='';

# ��ȡWord����
my $word = get_word();

foreach (@filelist) {
	next if /^\./; 
	print "����ת��: < ".$_." >Ϊ�ı��ļ�...\n";
	
	# Դ�ļ�
	$srcFile=$SRC_PATH.$_;
    #print $srcFile."\n";

	# �����е� .DOC/.doc �滻�� .txt
	s/.DOCX|.DOC$//i;#replace all .DOC

	# Ŀ���ļ�
	$destFile=$DEST_PATH."$_".'.txt';
	#print $destFile."\n";
	
	# ת���ļ���ʽ
	&convert($srcFile, $destFile);

}

# Word�ļ����ΪTXT�ļ�����Ҫ��������������
# 1. Դ�ļ�����·��
# 2. Ŀ���ļ�����·��
sub convert{
	my $oldFile=shift;
	my $newFile=shift;

	print $oldFile." -> ".$newFile."\n";

	# �趨��Wordʱ���ɼ�
	$word->{Visible} = 0;

	# ��ԭʼWord�ĵ�
	my $doc = $word->{Documents}->Open($oldFile);

	# ���ΪTXT�ļ�
	$doc->SaveAs(
	    $newFile,
	    wdFormatTextLineBreaks
	);

	# �ر�Word�ĵ�
	$doc->Close(0);

}

# ��ȡMicrosoft Word����
sub get_word {
    my $word;
    eval {
        $word = Win32::OLE->GetActiveObject('Word.Application');
    };

    die "$@\n" if $@;

    unless(defined $word) {
        $word = Win32::OLE->new('Word.Application', sub { $_[0]->Quit })
            or die "Oops, cannot start Word�� ",
                   Win32::OLE->LastError, "\n";
    }
    return $word;
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