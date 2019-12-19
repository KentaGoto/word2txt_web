use Mojolicious::Lite;
use File::Copy;
use open IO => qw/:encoding(UTF-8)/;
use HTML::Entities;
use XML::Twig;
use File::Find::Rule;
use File::Basename qw/basename dirname fileparse/;
use FindBin;
use Cwd;
use Archive::Zip qw( :ERROR_CODES :CONSTANTS :MISC_CONSTANTS );
use File::Path 'rmtree';
use File::Spec;
use Encode qw/encode decode/;

#binmode STDOUT, ':encoding(UTF-8)';

my $app = app;
my $url = 'http://localhost:3004'; # URL

my $UPLOAD_DIR = app->home->rel_file('/upload');
my $TMP_DIR = app->home->rel_file('/tmp');
my $PROC_TMP_DIR = 'proc_tmp';

my $nr;
my $ntoc;
my @texts;

get '/word2txt' => sub {
	my $c = shift;

	# キャッシュを残さない。
	# この設定がないと、ブラウザの戻るボタンでトップページに戻ってきた際に、選択したzipファイルを変更して実行しても以前選択したzipファイルで実行されてしまう。
	# FirefoxやEdgeではこの挙動となった。Chromeではこの挙動とはならなかった。2018/01/18
	$c->res->headers->header("Cache-Control" => "no-store, no-cache, must-revalidate, max-age=0, post-check=0, pre-check=0", 
							 "Pragma"        => "no-cache"
							);

	$c->render('index', top_page => $url);

	@texts = (); # リセット

	# uploadフォルダがなければ作成する
	if ( !-d $UPLOAD_DIR ){
		mkdir $UPLOAD_DIR or die "Can not create directory $UPLOAD_DIR";
	}
	# tmpフォルダがなかったら作る
	if ( -d $TMP_DIR ){
	
	} else {
		mkdir $TMP_DIR, 0700 or die "$!";
	}
	undef($c);
};

post '/word2txt' => sub {
    my $c = shift;
	my $nfn = $c->param('nfn');
	$nr = $c->param('nr');
	$ntoc = $c->param('ntoc');

    # 処理対象zipファイル
    my $file = $c->req->upload('files');
    my $zip_filename = $file->filename;
    
    # zip以外受け付けない
    unless ( $zip_filename =~ /\.zip$/ ){
    	return $c->render(
    		template => 'error', 
    		message  => "Error",
    		message2 => "Upload fail. The selected file is not an ZIP file.",
    	);
    }

	# Local time settings
	my $times = time();
	my ($sec, $min, $hour, $mday, $month, $year, $wday, $stime) = localtime($times);
	$month++;
	my $datetime = sprintf '%04d%02d%02d%02d%02d%02d', $year + 1900, $month, $mday, $hour, $min, $sec;

	# uploadディレクトリに日付フォルダ作成
	chdir $UPLOAD_DIR;
	if ( !-d $datetime ){
		mkdir $datetime or die "Can not create directory $datetime";
	}
    
    # アップロードされたファイルを保存
    my $upload_save_file = "$UPLOAD_DIR/" . "$datetime/" . $zip_filename;
    $file->move_to($upload_save_file);
    
	# tmpディレクトリに日付フォルダ作成
	chdir $TMP_DIR;
	if ( !-d $datetime ){
		mkdir $datetime or die "Can not create directory $datetime";
	}

	# tmpフォルダにも処理対象ファイルを移す
	my $tmp_save_file = "$TMP_DIR/" . "$datetime/" . $zip_filename;
	$file->move_to($tmp_save_file);
	
	# Zip解凍
	chdir $datetime;
	my $datetime_fullpath = "$TMP_DIR/$datetime";
	&unzip(\$zip_filename, $datetime_fullpath);

	# zip展開後はzipを削除
	unlink $tmp_save_file;

	# 処理ディレクトリとなるproc_tmpフォルダを作成
	mkdir $PROC_TMP_DIR or die "Can not create directory $PROC_TMP_DIR";
	
	# proc_tmpフォルダのフルパス
	my $PROC_TMP_DIR_abs = "$TMP_DIR/$datetime/$PROC_TMP_DIR";

	#########################################################################
	#  	docxからテキストを抽出						                     	 #
	#########################################################################
	my @docxs = File::Find::Rule->file->name( '*.docx', '*.DOCX' )->in(getcwd);
	chdir $PROC_TMP_DIR_abs;

	foreach (@docxs){
		my $docx_fullpath = $_;
		my $docx_filename = basename($docx_fullpath);
		my $docx_dirname  = dirname($docx_fullpath);

		print "Processing... $docx_filename\n";
	
		# [ファイル名区切りを出力する]がオンの場合
		if ( defined $nfn ){
			my $docx_filename_decode = decode('CP932', $docx_filename); # ファイル名をデコードしないとHTML出力で化ける
			push (@texts, "\n\n------------------------------$docx_filename_decode------------------------------");
		} else {
			# オフの場合は何もしない
		}
	
		# テキストボックスなどの文字列抽出処理
		&GetText ($docx_filename, $docx_dirname, $PROC_TMP_DIR_abs);
	}
	
	# resultsページに移り、抽出したテキストをリダイレクトする
	$c->redirect_to('/word2txt/results');
	
} => 'upload';

get '/word2txt/results' => sub {
    my $c = shift;
	@texts = grep $_ !~ /^\s*$/, @texts; # 空白のみまたは空は捨てる
	@texts = grep $_ !~ /PAGEREF _Toc/, @texts; # TOCは消す
	$c->render('results', 'texts' => \@texts);
};

sub GetText {
	my ($docx_filename, $docx_dirname, $PROC_TMP_DIR_abs) = @_;
	my $zip = &docxCopy2tmp($docx_filename, $docx_dirname, $PROC_TMP_DIR_abs); # docxをproc_tmpフォルダに移動してzipにする。
	&unzip(\$zip, $PROC_TMP_DIR_abs); # zip解凍
	unlink $zip; # 展開後のzipを削除
	unlink glob '*.xml'; # 要らないxmlを削除する
	my @xmls = File::Find::Rule->file->name( qr/(?:document\.xml|header\d*\.xml|footer\d*\.xml)/ )->in(getcwd);
	&xml_copy(\@xmls, $PROC_TMP_DIR_abs); # 対象のxmlファイルをproc_tmpフォルダにコピーする
	&del_dir($PROC_TMP_DIR_abs); # 要らないフォルダを削除
	unlink glob '*.rels'; # 要らないrelsファイルを削除
	my @target_xmls = glob '*.xml'; # proc_tmpフォルダにコピーしたxmlを対象とする
	&xml_parser(\@target_xmls); # xmlをパースしてテキストをゲットする
	unlink glob '*.xml'; # 対象ファイルの*.xmlを削除する
}

sub docxCopy2tmp {
	my ($docx_filename, $docx_dirname, $PROC_TMP_DIR_abs) = @_;
	my $zip;
	$zip = $docx_filename;
	$zip =~ s|^(.+)$|$1\.zip|;
	copy($docx_dirname . '/' . $docx_filename, "$PROC_TMP_DIR_abs/$zip") or die $!;
	return $zip;
}

sub unzip {
	my ($zip, $DIR) = @_;
	my $zip_obj = Archive::Zip->new($$zip);
	my @zip_members = $zip_obj->memberNames();
	foreach (@zip_members) {
		$zip_obj->extractMember($_, "$DIR/$_");
	}
}

sub xml_copy {
	my ($xmls, $PROC_TMP_DIR_abs) = @_;
	foreach (@$xmls){
		print $_ . "\n";
		my $file_src = $_;
		my $file_dst = $PROC_TMP_DIR_abs;
		copy($file_src, $file_dst) or die {$!};
	}
}

sub del_dir {
	my ($PROC_TMP_DIR_abs) = shift;
	rmtree("$PROC_TMP_DIR_abs/word") or die $!;
	rmtree("$PROC_TMP_DIR_abs/docProps") or die $!;
	rmtree("$PROC_TMP_DIR_abs/_rels") or die $!;
}

sub xml_parser {
	my ($target_xmls) = shift;
	foreach my $xml ( @$target_xmls ){
		my $twig = new XML::Twig( TwigRoots => {
				'//w:instrText' => \&fieldcode_element_delete, # フィールドコードを除外
				'//w:tr' => \&output_target,
				'//w:p' => \&output_target,
				});
		$twig->parsefile( $xml );
	}
}

sub output_target {
	my( $tree, $elem ) = @_;
	my $target = $elem->text;

	# [半角英数字を除外]がオンの場合
	if ( defined $nr ){
		if ( $target !~ m|\A[\x01-\x7E\xA1-\xDF\d\s\N{LEFT-TO-RIGHT MARK}\N{NEXT LINE}\N{RIGHT-TO-LEFT MARK}\N{LINE SEPARATOR}\N{PARAGRAPH SEPARATOR}™‰€“”–‘’…]+\z| ){
			push (@texts, $target);
		}
	} else {
		push (@texts, $target);
	}
	
	{
		local *STDOUT;
		local *STDERR;
  		open STDOUT, '>', undef;
  		open STDERR, '>', undef;
		$tree->flush_up_to( $elem ); #Memory clear
	}
}

sub fieldcode_element_delete {
	my ($twig, $element) = @_;
	my $target = $element->text;
	# [目次除外]がオンの場合
	if ( defined $ntoc ){
		# TOC以外のフィールドコード文字列を削除する。TOCは後で削除する。
		if ( $target !~ m|PAGEREF _Toc| ){
			$element->delete;
		}
	} else {
		# [目次除外]がオフの場合は、すべてのフィールドコード文字列を削除する
		$element->delete;
	}
}

app->start;

__DATA__

@@ error.html.ep
<h1><%= $message %></h1>
<p><%= $message2 %></p>

@@ layouts/default.html.ep
<html>
<head>
<title><%= title %></title>
<meta http-equiv="Content-type" content="text/html; charset=UTF-8">
<%= stylesheet '/css/style.css' %>
<link type="text/css" rel="stylesheet"
  href="http://code.jquery.com/ui/1.10.3/themes/cupertino/jquery-ui.min.css" />
<script type="text/javascript"
  src="http://code.jquery.com/jquery-1.10.2.min.js"></script>
<script type="text/javascript"
  src="http://code.jquery.com/ui/1.10.3/jquery-ui.min.js"></script>
</head>
<body><%= content %></body>
</html>

@@ index.html.ep
<%
	my $filename = stash('filename');
%>
% layout 'default';
% title 'word2txt';
%= javascript begin
  // プログレスバー
  $(document).on('click', '#run', function() {
    $('#progress').progressbar({
        max: 100,
        value: false
	});
	// ボタンなどの非表示
	// propやattrのグレーアウトは、Chromeだと処理が実行されないバグがあった。
	$('#run').hide(500);
	$('#select').hide(500);
	$('#delimiter_checkbox').hide(500);
	$('#not_required_checkbox').hide(500);
	$('#not_toc_checkbox').hide(500);
	$('#processing').show(500);
  });
% end
<div id="out">
<div id="head">
<h1>word2txt</h1>
<form method="post" action="<%= url_for('upload') %>" enctype ="multipart/form-data">
	<input name="files" type="file" id='select' value="Select File" />
	<input type="submit" id="run" value="Run" />
	</br>
	<p id="not_toc_checkbox"><strong>目次除外: </strong><%= check_box ntoc => 1, checked => "checked" %></p>
  	<p id="not_required_checkbox"><strong>半角英数字を除外: </strong><%= check_box nr => 1 %></p>
	<p id="delimiter_checkbox"><strong>ファイル名区切りを出力: </strong><%= check_box nfn => 1 %></p>
	</br>
	<p id="processing" style="display: none;">Processing... </p>
	<div id="progress"></div>
</form>
	</div>
	<div id="main">
<h3>Usage</h3>
	<ul>
		<li><strong>*.docx</strong> ファイルの入った <strong>zip</strong> ファイルを選択します。</li>
		<li><strong>[Run]</strong> ボタンをクリックします。</li>
		<li>遷移した画面に <strong>*.docx</strong> から抽出したテキストが表示されます。</li>
	</ul>
<h3>Option</h3>
	<ul>
		<li><strong>[目次除外]</strong> チェックボックスはデフォルトでオンです。目次を除外したくない場合はオフにしてください。</li>
		<li><strong>[半角英数字を除外]</strong> チェックボックスは、英訳の際に半角英数字を除外したい場合にオンにしてください。</br>※制御文字やいつくかの特殊文字も対象になります。</li>
		<li><strong>[ファイル名区切りを出力]</strong> チェックボックスをオンにすると、処理対象となった <strong>*.docx</strong> が抽出テキストの区切りとして出力されます。</li>
	</ul>
<h3>Requirements</h3>
	<ul>
		<li>Chrome or Firefox</li>
	</ul>
<h4>Note</h4>
<ul>
	<li>本文、ヘッダー、フッター、テキストボックスに使用されている文字が抽出されます。</li>
</ul>
</div>
<div id="footer">
Copyright &copy; KentaGoto All Rights Reserved.
</div>
</div>

@@ results.html.ep
<html>
<head>
% title 'Results';
<meta http-equiv="Content-type" content="text/html; charset=UTF-8">
<%= stylesheet '/css/style_Results.css' %>
</head>
<body>
% for my $t (@$texts){
	<%= $t %></br>
% }
</body>
</html>
