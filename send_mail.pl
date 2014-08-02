printf("Iniciando o envio de e-mails... \n");
use strict;
use Try::Tiny;
use IO::All;
use Email::MIME;
use Email::Sender::Simple qw(sendmail);
use Email::Sender::Transport::SMTP::TLS;
use Spreadsheet::XLSX;
use File::Basename;
use Module::Runtime;

my $diretorioScript = dirname($0);
my $excel = Spreadsheet::XLSX->new( "$ENV{HOMEDRIVE}" . "\\JOFA\\jofalista.xlsx", );

my $diretorio = "$diretorioScript\\arquivos\\";
my $email = "";
my $nome = "";


#CONF MAIL
my $login = '';
my $senha = "";

opendir(diretorio, "$diretorio");
my @lista = readdir(diretorio);
closedir(diretorio);

foreach my $arquivo(@lista) {
	$email = "";
	$nome = "";
	
	if ($arquivo ne "." && $arquivo ne ".."){
						
		# MINERANDO O XLSX PROCURANDO PELO NOME DO ARQUIVO NA PRIMEIRA COLUNA
		# COLUNA 0 - NOME DA PESSOA - MESMO NOME DO ARQUIVO
		# COLUNA 1 - E-MAIL QUE SERÁ ENVIADO
		
		foreach my $sheet ( @{ $excel->{Worksheet} } ) {
			$sheet->{MaxRow} ||= $sheet->{MinRow};
			
			# LOOP LINHAS
			foreach my $row ( $sheet->{MinRow} .. $sheet->{MaxRow} ) {
				$sheet->{MaxCol} ||= $sheet->{MinCol};
				
				# LOOP COLUNAS
				foreach my $col ( $sheet->{MinCol} .. $sheet->{MaxCol} ) {
					
					my $cell = $sheet->{Cells}[$row][$col];					
					if ($cell) {											
						if("$cell->{Val}.pdf" eq $arquivo){
							$nome = $cell->{Val};
							my $emailCell = $sheet->{Cells}[$row][$col+1];
							
							# ATRIBUINDO O EMAIL ENCONTRADO
							$email = $emailCell->{Val};							
						}												
					}    
				}
			}
		}
		
		#VERIFICA SE O CAMPO DE E-MAIL NÃO ESTÁ EM BRANCO 
		if($email ne ""){
			printf("Processando o arquivo $arquivo para o email: $email \n");
			
			# CRIANDO UM ARRAY COM ANEXOS
			my @parts = (
				Email::MIME->create(
					attributes => {
						filename     => "$arquivo",
						content_type => "application/pdf",
						encoding     => "base64",
						disposition  => "attachment",
						name         => "$arquivo",
					},
					body => io( "$diretorio$arquivo" )->all,
				),
				
				# CRIANDO O CONTEÚDO DO E-MAIL COM ANEXOS
				Email::MIME->create(
					attributes => {
						content_type  => "text/html",
					},
					body => "Ola $nome, segue em anexo o certificado dos cursos realizados na JOFA 2014.",
				)
			);
 			
			my $email_object = Email::MIME->create(
				header => [
					From           => '',
					To             => $email,
					Subject        => "",
					content_type   =>'multipart/mixed'
				],
				parts  => [ @parts ],
			);
 
			#FAZENDO A CONEXÃO SMTP GMAIL
			my $transport = Email::Sender::Transport::SMTP::TLS->new(
				host     => 'smtp.gmail.com',
				port     => 587,
				username => $login,
				password => $senha,
				timeout  => 500
			);
 
			# send the mail
			try {
				   sendmail( $email_object, {transport => $transport} );
			} catch {
				   warn "Email sending failed: $_";
			};
		}
	}

}

printf("Pressione ENTER para Finalizar... \n");
chomp( my $input = <STDIN> );