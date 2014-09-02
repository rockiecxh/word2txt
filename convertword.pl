#!/usr/bin/perl -w
# usage: perl conword.pl srcFolder destFolder

use strict;
use warnings;
use Cwd;
use File::Spec::Functions qw( catfile );

use Win32::OLE;
use Win32::OLE::Const 'Microsoft Word';

$Win32::OLE::Warn = 3;


# 获取当前目录的完整路径信息。
my $currentDir = getcwd;
print "当前目录: ".$currentDir."\n";

# 设置源文件夹
my $SRC_PATH = "$currentDir/Word/";

# 设置目标文件夹
my $DEST_PATH="$currentDir/TXT/";

deldir($DEST_PATH) if( -d $DEST_PATH);

# 如果目标文件夹不存在，则创建文件夹
mkdir( $DEST_PATH, 0777 ) if ( !-d $DEST_PATH);

opendir TEMP, ${SRC_PATH} or die "无法打开".$SRC_PATH."目录，请检查一下目录是否存在！";

# 读取目录下所有文件
my @filelist = readdir TEMP; 

my $srcFile='';
my $destFile='';

# 获取Word程序
my $word = get_word();

foreach (@filelist) {
	next if /^\./; 
	print "正在转换: < ".$_." >为文本文件...\n";
	
	# 源文件
	$srcFile=$SRC_PATH.$_;
    #print $srcFile."\n";

	# 将所有的 .DOC/.doc 替换成 .txt
	s/.DOCX|.DOC$//i;#replace all .DOC

	# 目标文件
	$destFile=$DEST_PATH."$_".'.txt';
	#print $destFile."\n";
	
	# 转换文件格式
	&convert($srcFile, $destFile);

}

# Word文件另存为TXT文件，需要传入两个参数：
# 1. 源文件完整路径
# 2. 目标文件完整路径
sub convert{
	my $oldFile=shift;
	my $newFile=shift;

	print $oldFile." -> ".$newFile."\n";

	# 设定打开Word时不可见
	$word->{Visible} = 0;

	# 打开原始Word文档
	my $doc = $word->{Documents}->Open($oldFile);

	# 另存为TXT文件
	$doc->SaveAs(
	    $newFile,
	    wdFormatTextLineBreaks
	);

	# 关闭Word文档
	$doc->Close(0);

}

# 获取Microsoft Word程序
sub get_word {
    my $word;
    eval {
        $word = Win32::OLE->GetActiveObject('Word.Application');
    };

    die "$@\n" if $@;

    unless(defined $word) {
        $word = Win32::OLE->new('Word.Application', sub { $_[0]->Quit })
            or die "Oops, cannot start Word！ ",
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