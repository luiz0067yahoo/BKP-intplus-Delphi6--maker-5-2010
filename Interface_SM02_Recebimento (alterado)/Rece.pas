unit Rece;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, jpeg, ExtCtrls, StdCtrls, Buttons, ComCtrls, Gauges, Spin;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    Image1: TImage;
    StatusBar1: TStatusBar;
    Memo1: TMemo;
    BitBtn1: TBitBtn;
    BitBtn4: TBitBtn;
    GroupBox1: TGroupBox;
    Gauge1: TGauge;
    Gauge3: TGauge;
    Gauge4: TGauge;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    GroupBox2: TGroupBox;
    Gauge2: TGauge;
    Timer1: TTimer;
    Timer2: TTimer;
    Timer3: TTimer;
    BitBtn2: TBitBtn;
    Timer4: TTimer;
    CheckBox1: TCheckBox;
    SpinEdit1: TSpinEdit;
    procedure LerArquivo(NomeDoArquivo: String );
    function  LerLinha(File_Name: String; Line: Integer): String;
    procedure SalvaLinha(File_Name: String; Text: String; Line: Integer);
    procedure INICIALIZA_PORTA(NR_PORTA: String; Velocidade_Porta: Integer; Paridade: String; NR_Bits: Integer; StopBits: Integer);
    procedure TRANSMITE_BLOCO(COMANDO_TX: Integer; NR_DADOS_BLOCO: Integer);
    procedure Salva_Linha_Batida(Endereco_HUM: Integer; LINHA_A_SALVAR: Integer);
    procedure CARREGA_HORARIO;
    procedure OpenURL(Url: String);
    procedure Image1Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure Timer3Timer(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);

		function Gerar_Nome_Arquivo_Registros(): String;
    procedure Timer4Timer(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
    procedure SpinEdit1Exit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
    Apagar																				: Integer;
    hCommFile																			: THandle;
    ForamTx																			 	: DWord;
    DadoRx																				: Array [1..250 ] of Byte;
    DadoTx																			 	: Array [1..250 ] of Byte;
    CRC																						: Integer;
    CRC_COMPARA 																	: Integer;
    vetor																			 		: Integer;
    Nr_Bytes																			: Integer;
    Comando																			 	: Integer;
    TX_Ok 																				: Boolean;
    Tx_Horario																		: Boolean;
    REG_FUNC_SALVAR																: Array [1..12] of Byte;
    REG_VETOR																			: Integer;
    VETOR_FUNC																		: Integer;
    NR_LINHAS_ARQUIVO															: Integer;
    S        																			: TextFile;
    Linha																					: String;
    S1 																						: String;
    Texto    																		 	: Array [ 1..25585 ] of String;
    MENSAGEM_TELA																	: Array [1..53] of String ;
    S_TX																					: AnsiString;
    S_TX2																				 	: AnsiString;
    DATA_BATIDA																		: AnsiString;
    HORA_BATIDA																		: AnsiString;
    NR_CARTAO																			: AnsiString;
    TIPO_BATIDA																		: AnsiString;
    Capacidade_Memoria													 	: Integer;
    Conectado																			: Boolean;
    Transacao 																	 	: Boolean;
    Numero_Do_Ponto 															: Integer;
    Habilita_Clear 																: Boolean;
    Quantidade_Batidas														: Integer;
    caminhoini																	 	: String;
    caminhobatidas																: String;
    caminhoshorarios															: String;
    caminhospessoas															 	: String;
  end;

    const BATIDAS               									= 60;
    const HORARIOS              									= 70;
    const FUNCIONARIOS          									= 80;
    const DATA_HORA             									= 90;
    const PEGA_INDICE_FUNC      									= 100;
    const PEGA_INDICE_HOR       									= 101;
    const PEGA_INDICE_BATI      									= 102;
    const LISTA_FULL            									= 105;
    const LISTA_FULL_INI        									= 106;
    const PING                  									= 50;
    const RETRANSMISSAO         									= 40;
    const APAGAR_FUNC_CONF      									= 30;
    const APAGAR_HOR_CONF       									= 32;
    const APAGAR_BATIDAS_CONF   									= 34;
    const APAGAR_FUNC           									= 31;
    const APAGAR_HOR            									= 33;
    const APAGAR_BATIDAS        									= 35;
    const LISTA_FUNC_INI        									= 20;
    const LISTA_FUNC            									= 21;
    const COM_AJUSTA_RTC        									= 110;
    const LISTA_HOR	        											= 120;
    const LISTA_HOR_INI	        									= 121;
    const MUDA_MENSAGEM         									= 130;

var
  Form1: TForm1;



implementation

{$R *.dfm}

uses ShellAPI, DateUtils;


{-----------------------------------------------------------------------------------------------------------------------}


function TForm1.Gerar_Nome_Arquivo_Registros(): String;
var
	vDia_S											: String;
  vMes_S											: String;
  vAno_S											: String;
  vHoras_S										: String;
  vMinutos_S									: String;
  vSegundos_S									: String;
begin

	vDia_S			:= FloatToStr( DayOf(Now()) );
  While (Length(vDia_S) < 2) Do vDia_S := ('0' + vDia_S);

  vMes_S			:= FloatToStr( MonthOf(Now()) );
  While (Length(vMes_S) < 2) Do vMes_S := ('0' + vMes_S);

  vAno_S			:= FloatToStr( YearOf(Now()) );
  While (Length(vAno_S) < 2) Do vAno_S := ('0' + vAno_S);

  vHoras_S		:= FloatToStr( HourOf(Now()) );
  While (Length(vHoras_S) < 2) Do vHoras_S := ('0' + vHoras_S);

  vMinutos_S	:= FloatToStr( MinuteOf(Now()) );
  While (Length(vMinutos_S) < 2) Do vMinutos_S := ('0' + vMinutos_S);

  vSegundos_S	:= FloatToStr( SecondOf(Now()) );
  While (Length(vSegundos_S) < 2) Do vSegundos_S := ('0' + vSegundos_S);

//  Result := ExtractFilePath( Application.ExeName ) + 'Registros_' + vDia_S + vMes_S + vAno_S + '_' + vHoras_S + vMinutos_S + vSegundos_S + '.txt';

  Result := 'C:\Registros_' + vDia_S + vMes_S + vAno_S + '_' + vHoras_S + vMinutos_S + vSegundos_S + '.txt';

end;
{-----------------------------------------------------------------------------------------------------------------------}


procedure TForm1.LerArquivo(NomeDoArquivo: String);
begin

	If (NOT FileExists(NomeDoArquivo)) then
		begin
			Memo1.Lines.Add(TimeToStr(Time) + ' Arquivo não existe!');
			NR_LINHAS_ARQUIVO := 1;
      Exit;
		end;


  AssignFile(S, NomeDoArquivo);
  Reset(S);
  Vetor := 1;
  ReadLn(S, Linha);

  TEXTO[vetor] := Linha;

  While (NOT Eof(S)) Do
  	begin
  		ReadLn(S, Linha);
  		Vetor := (Vetor + 1);
  		TEXTO[vetor] := Linha;
    end;

  NR_LINHAS_ARQUIVO := VETOR;
  CloseFile(S);

end;
{-----------------------------------------------------------------------------------------------------------------------}


function  TForm1.LerLinha(File_Name: String; Line: Integer): String;
begin

	LerArquivo(File_Name);

	result := TEXTO[Line];

end;
{-----------------------------------------------------------------------------------------------------------------------}


procedure TForm1.SalvaLinha(File_Name: String; Text: AnsiString; Line: Integer);
begin

	LerArquivo(File_Name);
 	TEXTO[Line] := Text;

  AssignFile(S, File_Name);
  Rewrite(S);

  vetor := 1;
  While (vetor <> (LINE + 1)) Do
  	begin
  		Write(S, TEXTO[vetor]);
  		WriteLn(S);
  		Vetor := (Vetor + 1);
  	end;

  CloseFile(S);

end;
{-----------------------------------------------------------------------------------------------------------------------}


procedure TFORM1.INICIALIZA_PORTA(NR_PORTA:string; Velocidade_Porta:integer ;Paridade:string; NR_Bits:integer; StopBits:integer ) ;
Var
 config												: String;
 DCB													: TDCB;
 CommPort											: String;
 Calibra_Timeouts							: TCommTimeouts;
begin

	CommPort 	:= NR_PORTA;
  hCommFile	:= CreateFile(Pchar(CommPort), GENERIC_READ OR GENERIC_WRITE, 0, nil, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
	//dcb.Flags := dcb.Flags or dcb_OutxCtsFlow or dcb_RtsControlHandshake;

	If (hCommFile = INVALID_HANDLE_VALUE) then
  	begin
      Application.MessageBox(PChar('Não foi possível abrir a ' + CommPort), 'ATENÇÃO', MB_OK + MB_ICONWARNING);
      Exit;
    end;

	If (NOT GetCommState(hCommFile,DCB)) then
  	begin
    	Application.MessageBox(PChar('Erro executando o GetCommState da ' + CommPort), 'ATENÇÃO', MB_OK + MB_ICONWARNING);
      Exit;
    end;

	Config := NR_PORTA + ': baud=' + IntToStr(Velocidade_Porta) + ' parity=' + Paridade + ' data=' + IntToStr(NR_Bits) + ' stop=' + IntToStr(StopBits);

	If (NOT BuildCommDCB(@Config[1], DCB)) then
  	begin
    	Application.MessageBox(PChar('Erro construindo DCB para ' + CommPort), 'ATENÇÃO', MB_OK + MB_ICONWARNING);
      Exit;
  	end;

	If (NOT SetCommState(hCommFile, DCB)) then
  	begin
    	Application.MessageBox(PChar('Erro executando o SetCommState para ' + CommPort), 'ATENÇÃO', MB_OK + MB_ICONWARNING);
      Exit;
    end;

  Calibra_Timeouts.ReadIntervalTimeout 					:= 100;
  Calibra_Timeouts.ReadTotalTimeoutMultiplier 	:= 1;
  Calibra_Timeouts.ReadTotalTimeoutConstant 	 	:= 100;
  Calibra_Timeouts.WriteTotalTimeoutMultiplier	:= 1000;
  Calibra_Timeouts.WriteTotalTimeoutConstant 		:= 1000;

	If (NOT SetCommTimeouts(hCommFile, Calibra_Timeouts)) then
  	begin
    	Application.MessageBox(PChar('Erro executando SetCommTimeouts ' + CommPort), 'ATENÇÃO', MB_OK + MB_ICONWARNING);
      Exit;
    end;

	ForamTx := 3;

end;
{-----------------------------------------------------------------------------------------------------------------------}


procedure TForm1.TRANSMITE_BLOCO(COMANDO_TX: Integer; NR_DADOS_BLOCO: Integer);
begin

  DadoTx[1]:= 170;
  DadoTx[2]:= 21;
  DadoTx[3]:= 85;
  DadoTx[4]:= 128;
  DadoTx[5]:= 1;
  DadoTx[6]:= (2 + NR_DADOS_BLOCO);
  DadoTx[7]:= (COMANDO_TX Div 256);
  DadoTx[8]:= (COMANDO_TX And $FF);

  vetor := 9;
  CRC   := (DadoTx[7] + DadoTx[8]);

  While (vetor <> (NR_DADOS_BLOCO + 9)) Do
  	begin
    	CRC 	:= (CRC + DADOTX[vetor]);
    	vetor	:= (vetor + 1);
   	end;

	DadoTx[vetor]			:= (CRC Div 256);
  DadoTx[vetor + 1]	:= (CRC And $00FF);

  Vetor := (Vetor + 1);

  WriteFile(hCommFile, DadoTx, vetor, ForamTx, nil);

end;
{-----------------------------------------------------------------------------------------------------------------------}


procedure TFORM1.OpenURL(Url: string);
begin
  ShellExecute (0, 'open', PChar(url), PChar(url), nil, SW_SHOW);
end;
{-----------------------------------------------------------------------------------------------------------------------}


procedure TForm1.Salva_Linha_Batida(Endereco_HUM: Integer; LINHA_A_SALVAR: Integer);
var
	MES										 			: String;
  ANO										 			: String;
  NR_POINT									 	: String;
  HORA										 		: String;
  MINUTO										 	: String;
begin

  DATA_BATIDA := IntToStr(DADORX[6 + Endereco_HUM] And 31) ;

  If (Length(DATA_BATIDA) < 2) then DATA_BATIDA := ('0' + DATA_BATIDA);
  MES := IntToStr(DADORX[9 + Endereco_HUM] Div 16);

  If (Length(MES) = 1 ) then MES := ('0' + MES);
  ANO := IntToStr(DADORX[9 + Endereco_HUM] And $0f);

  If (Length(ANO) = 1) then ANO := ('0' + ANO);

  DATA_BATIDA := (DATA_BATIDA + '/' + MES + '/' + ANO);
  HORA   := IntToStr(DADORX[7 + Endereco_HUM]);
  Minuto := IntToStr(DADORX[8 + Endereco_HUM]);

  If (Length(HORA) = 1) then HORA := ('0' + HORA);

  If (Length(Minuto) = 1) then Minuto := ('0' + Minuto);
  HORA_BATIDA := (HORA + ':' + MINUTO);

  If ((DADORX[6 + Endereco_HUM] And 32) = 32) then TIPO_BATIDA := 'E';

  If ((DADORX[6 + Endereco_HUM] And 64) = 64) then TIPO_BATIDA := 'S';

  NR_CARTAO := IntToStr((DADORX[3 + Endereco_HUM] * 65535) + (DADORX[4 + Endereco_HUM] * 256) + DADORX[5 + Endereco_HUM]);

  While (Length(NR_CARTAO) <> 8) Do
  	begin
  		NR_CARTAO := ('0' + NR_CARTAO);
  	end;

	NR_POINT := IntToStr(Numero_Do_Ponto);

  If (Length(NR_POINT) = 1) then NR_POINT := ('0' + NR_POINT);

  S_TX  := (DATA_BATIDA + ' ' + HORA_BATIDA + ' ' + '000000000000' + NR_CARTAO + ' ' + NR_POINT + ' ' + TIPO_BATIDA + ' 00');

  Texto[VETOR_FUNC] := (S_TX + #13 + #10);

end;
{-----------------------------------------------------------------------------------------------------------------------}


procedure TForm1.CARREGA_HORARIO ;
var
	VETOR_STR										: Integer;
  SEMANA_INTEIRA							: Integer;
begin

	S_TX  := LerLinha(CaminhosHorarios, VETOR_FUNC);
	DADOTX[9] 	:= 0;
	DADOTX[10] 	:= StrToInt(S_TX[1] + S_TX[2]);

  SEMANA_INTEIRA := 1;

  While (SEMANA_INTEIRA < 8) Do
  	begin

    	S_TX  := LerLinha(CaminhosHorarios, VETOR_FUNC);

    	While (Length(S_TX) <> 63) Do S_TX := (S_TX + '9');

    	REG_VETOR  := ((StrToInt(S_TX[3]) - 1) * 24);

    	VETOR_STR := 0;
    	While (Vetor_Str < 60) Do
    		begin
          // 1 expediente
          DADOTX[11 + REG_VETOR] := StrToInt(S_TX[4  + VETOR_STR] + S_TX[5  + VETOR_STR]);
          DADOTX[12 + REG_VETOR] := StrToInt(S_TX[7  + VETOR_STR] + S_TX[8  + VETOR_STR]);
          DADOTX[13 + REG_VETOR] := StrToInt(S_TX[9  + VETOR_STR] + S_TX[10 + VETOR_STR]);
          DADOTX[14 + REG_VETOR] := StrToInt(S_TX[12 + VETOR_STR] + S_TX[13 + VETOR_STR]);
          //
          REG_VETOR := (REG_VETOR + 4);
          Vetor_Str := (Vetor_Str + 10);
    		end;

    	VETOR_FUNC 			:= (VETOR_FUNC + 1);
    	SEMANA_INTEIRA 	:= (SEMANA_INTEIRA + 1);
    end;

end;
{-----------------------------------------------------------------------------------------------------------------------}


procedure TForm1.Image1Click(Sender: TObject);
begin
	OpenURL('http://www.smartbr.com');
end;
{-----------------------------------------------------------------------------------------------------------------------}


procedure TForm1.BitBtn1Click(Sender: TObject);
begin

  if (not CheckBox1.Checked) then
  begin
    If (QUANTIDADE_BATIDAS = 0) then
      begin
        Application.MessageBox('Não há registros na memória... Operação cancelada!', 'Operação Cancelada', MB_OK + MB_ICONINFORMATION);
        Exit;
      end;

    If (Application.MessageBox('Deseja realmente receber os dados referentes ao ponto?', 'Receber Dados', MB_YESNO + MB_ICONQUESTION) <> ID_YES) then Exit;

    BitBtn1.Enabled	:= False;

    gauge2.Progress	:= 0;

    VETOR_FUNC := 1;

    Timer2.Enabled	:= False;
    Timer3.Enabled	:= False;

    TRANSMITE_BLOCO(LISTA_FULL_INI, 0);
  end else
    Application.MessageBox('Desabilite a opção de Modo Automático!', 'Operação não Realizada', MB_OK + MB_ICONINFORMATION);
end;
{-----------------------------------------------------------------------------------------------------------------------}


procedure TForm1.BitBtn4Click(Sender: TObject);
begin
 Close ;
end;
{-----------------------------------------------------------------------------------------------------------------------}


procedure TForm1.FormCreate(Sender: TObject);
var automatico: boolean;
begin

 CaminhoIni				:= ExtractFilePath(Application.ExeName) + '\config.ini';
 CaminhoBatidas		:= ExtractFilePath(Application.ExeName) + '\batidas.txt';
 CaminhosHorarios	:= ExtractFilePath(Application.ExeName) + '\horarios.txt';
 CaminhosPessoas	:= ExtractFilePath(Application.ExeName) + '\pessoas.txt';

	If (NOT FileExists(CaminhoIni)) then
  	begin
    	Memo1.Clear;
    	Memo1.Lines.SaveToFile(CaminhoIni);
    	SalvaLinha(CaminhoIni, 'COM1', 1);
    	SalvaLinha(CaminhoIni, '0', 2);
    	SalvaLinha(CaminhoIni, '60', 3);
      SalvaLinha(CaminhoIni, 'true', 4);
    end;

  try
    if (LerLinha(caminhoini, 4) = 'true') then
      automatico := true
    else begin
      SalvaLinha(CaminhoIni, 'false', 4);
      automatico := false;
    end;  
  Except
      SalvaLinha(CaminhoIni, 'false', 4);
      automatico := false;
  end;

  try
    Timer4.Interval := StrToInt(LerLinha(caminhoini, 3)) * 1000;
  Except
    Timer4.Interval := 60000;
    SalvaLinha(CaminhoIni, '60', 3);
  end;

  //if ( not FileExists (caminhobatidas) ) then begin
  Memo1.Clear;
	Memo1.Lines.SaveToFile(CaminhoBatidas);
  //end;

  Numero_Do_Ponto := StrToInt(LerLinha(CaminhoIni, 2));

	INICIALIZA_PORTA(LerLinha (CaminhoIni, 1), 19200, 'n', 8, 2);
  Apagar := 0;
  Tx_Horario := False  ;
  VETOR_FUNC := 1 ;

  TRANSMITE_BLOCO(PING, 0);

  Conectado := True;
  Transacao := False;
  Habilita_Clear := False;

  Timer4.Enabled    := automatico;
  CheckBox1.Checked := automatico;
  SpinEdit1.Value   := Timer4.Interval div 1000;

end;
{-----------------------------------------------------------------------------------------------------------------------}


procedure TForm1.Timer1Timer(Sender: TObject);
var
	Aviso												: Boolean;
	vLinha											: String;
	vRegistro										: String;
  vRegistros									: TStringList;
begin

	vetor := 1;
	While (vetor <> 251) Do
  	begin
  		DadoRx[vetor] := 0;
  		vetor := (vetor + 1);
 		end;

	ReadFile(hCommFile, DadoRx, 6, ForamTx, nil);

	If ((DadoRx[1] = 170) And (DadoRx[4] = 128) And (DadoRx[5] = 1)) then
		begin

      Nr_Bytes := DadoRx [6] + 2 ;

      vetor := 1 ;
      While (vetor <> 10 ) Do
      	begin
      		DadoRx[vetor] := 0;
      		vetor := (vetor + 1);
      	end;

      ReadFile(hCommFile, DadoRx, Nr_Bytes, ForamTx, nil);
      Comando := ( (DadoRx[1] * 256) + DadoRx[2]);

      CRC := 0;

      vetor := 1;
      While (vetor <> (Nr_Bytes - 1)) Do
      	begin
      		CRC   := (CRC + DadoRx[vetor]);
      		vetor := (vetor + 1);
        end;
      CRC_COMPARA := (DadoRx[vetor ] * 256);
      CRC_COMPARA := (CRC_COMPARA + DadoRx[vetor + 1]);

      If (CRC = CRC_COMPARA) then
				begin
 					Timer3.Enabled := False;
 					Timer3.Enabled := True;

					Case (Comando) Of
						PING             										 	: begin

                                                      If (NOT Transacao) then BITBTN1.Enabled := True;
                                                      TRANSMITE_BLOCO(PEGA_INDICE_HOR, 0);

                                                      StatusBar1.Panels[0].Text := ('Status: Conectado a 19200bps na ' + LerLinha(CaminhoIni, 1));
                                                      If (Conectado) then
                                                      	begin
                                                      		MEMO1.Lines.Add(TimeToStr(TIME) + ' Conectado');
                                                      		Conectado := False;
                                                        end;

            																				end;

						PEGA_INDICE_HOR  											: begin

                                                      Capacidade_Memoria := ((DADORX[3] * 65536) + (DADORX[4] * 256) + DADORX[5]);
                                                      gauge1.Progress := ((100 * Capacidade_Memoria) Div 340);
                                                      TRANSMITE_BLOCO(PEGA_INDICE_FUNC, 0);
            																				end;

						PEGA_INDICE_FUNC 											: begin
                                                      Capacidade_Memoria 	:= ((DADORX[3] * 65536) + (DADORX[4] * 256) + DADORX[5]);
                                                      gauge3.Progress			:= ((100 * Capacidade_Memoria) Div 500);
                                                      TRANSMITE_BLOCO(PEGA_INDICE_BATI, 0);
            																				end;

            PEGA_INDICE_BATI											: begin
                                                      Capacidade_Memoria 	:= ((DADORX[3] * 256) + DADORX[4]);
                                                      Quantidade_Batidas 	:= Capacidade_Memoria;
                                                      gauge4.Progress			:= ((100 * Capacidade_Memoria) Div 24585);
                                                      Gauge2.MaxValue			:= ((DADORX[3] * 256) + DADORX[4]);
                                                    end;

						LISTA_FULL       											: begin

                     																	If (NOT ((DADORX[3] = 0) And (DADORX[4] = 0) And (DADORX[5] = 0))) then
                                                      	begin
                                                        	Salva_Linha_Batida(0, VETOR_FUNC);
                                                        	VETOR_FUNC 			:= (VETOR_FUNC + 1);
                                                        	Gauge2.Progress := (Gauge2.Progress + 1);
                                                    		end;

                                                      If (NOT ((DADORX[10] = 0) And (DADORX[11] = 0) And (DADORX[12] = 0))) then
                                                      	begin
                                                      		Salva_Linha_Batida(7, VETOR_FUNC);
                                                      		VETOR_FUNC 			:= (VETOR_FUNC + 1);
                                                      		Gauge2.Progress	:= (Gauge2.Progress + 1);
                                                      	end;

                                                      If (NOT ((DADORX[17] = 0) And (DADORX[18] = 0) And (DADORX[19] = 0))) then
                                                      	begin
                                                      		Salva_Linha_Batida(14, VETOR_FUNC);
                                                      		VETOR_FUNC 			:= (VETOR_FUNC + 1);
                                                      		Gauge2.Progress	:= (Gauge2.Progress + 1);
                                                        end;

																											If (NOT ((DADORX[24] = 0) And (DADORX[25] = 0) And (DADORX[26] = 0))) then
                                                      	begin
                                                      		Salva_Linha_Batida(21, VETOR_FUNC);
                                                      		VETOR_FUNC 			:= (VETOR_FUNC + 1);
                                                      		Gauge2.Progress	:= (Gauge2.Progress + 1);
                                                      	end;

                                                     	If (NOT ((DADORX[31] = 0) And (DADORX[32] = 0) And (DADORX[33] = 0))) then
                                                      	begin
                                                      		Salva_Linha_Batida(28, VETOR_FUNC);
                                                          VETOR_FUNC 			:= (VETOR_FUNC + 1);
                                                          Gauge2.Progress := (Gauge2.Progress + 1);
                                                      		MEMO1.Lines.Add(TimeToStr(TIME) + ' Registro ' + IntToStr(VETOR_FUNC-1));
                                                     		end;

                                                      If (NOT ((DADORX[38] = 0) And (DADORX[39] = 0) And (DADORX[40] = 0))) then
                                                      	begin
                                                      		Salva_Linha_Batida(35, VETOR_FUNC);
                                                      		VETOR_FUNC 			:= (VETOR_FUNC + 1);
                                                      		Gauge2.Progress := (Gauge2.Progress + 1);
                                                      	end;

                                                      If (NOT ((DADORX[45] = 0) And (DADORX[46] = 0) And (DADORX[47] = 0))) then
                                                      	begin
                                                      		Salva_Linha_Batida(42, VETOR_FUNC);
                                                      		VETOR_FUNC 			:= (VETOR_FUNC + 1);
                                                      		Gauge2.Progress	:= (Gauge2.Progress + 1);
                                                        end;

                                                      If (NOT ((DADORX[52] = 0) And (DADORX[53] = 0) And (DADORX[54] = 0))) then
                                                      	begin
                                                      		Salva_Linha_Batida(49, VETOR_FUNC);
                                                      		VETOR_FUNC 			:= (VETOR_FUNC + 1);
                                                      		Gauge2.Progress := (Gauge2.Progress + 1);
                                                      	end;

                                                      If (NOT ((DADORX[59] = 0) And (DADORX[60] = 0) And (DADORX[61] = 0))) then
                                                      	begin
                                                      		Salva_Linha_Batida(56, VETOR_FUNC);
                                                      		VETOR_FUNC 			:= (VETOR_FUNC + 1);
                                                      		Gauge2.Progress	:= (Gauge2.Progress + 1);
                                                      	end;

                                                      If (NOT ((DADORX[66] = 0) And (DADORX[67] = 0) And (DADORX[68] = 0))) then
                                                      	begin
                                                      		Salva_Linha_Batida(63, VETOR_FUNC);
                                                      		VETOR_FUNC 			:= (VETOR_FUNC + 1);
                                                      		Gauge2.Progress	:= (Gauge2.Progress + 1);
                                                      		MEMO1.Lines.Add(TimeToStr(TIME) + ' Registro '+IntToStr(VETOR_FUNC - 1));
                                                      	end;

                                                     	Aviso := True;
{
                                                      If (NOT ((DADORX[3] = 0) And (DADORX[4] = 0) And (DADORX[5] = 0))) then
                                                      	begin
                                                      		If (NOT ((DADORX[10] = 0) And (DADORX[11] = 0) And (DADORX[12] = 0))) then
                                                        		begin
                                                      				If (NOT ((DADORX[17] = 0) And (DADORX[18] = 0) And (DADORX[19] = 0))) then
                                                              	begin
                                                      						If (NOT ((DADORX[24] = 0) And (DADORX[25] = 0) And (DADORX[26] = 0))) then
                                                                  	begin
                                                      								If (NOT ((DADORX[31] = 0) And (DADORX[32] = 0) And (DADORX[33] = 0))) then
                                                                      	begin
                                                                        	If (NOT ((DADORX[38] = 0) And (DADORX[39] = 0) And (DADORX[40] = 0))) then
                                                                          	begin
                                                                            	If (NOT ((DADORX[45] = 0) And (DADORX[46] = 0) And (DADORX[47] = 0))) then
                                                                              	begin
                                                                              		If (NOT ((DADORX[52] = 0) And (DADORX[53] = 0) And (DADORX[54] = 0))) then
                                                                                  	begin
                                                                              				If (NOT ((DADORX[59] = 0) And (DADORX[60] = 0) And (DADORX[61] = 0))) then
                                                                                      	begin
                                                                              						If (NOT ((DADORX[66] = 0) And (DADORX[67] = 0) And (DADORX[68] = 0))) then
                                                                                          	begin
                                                                              								TRANSMITE_BLOCO ( LISTA_FULL,0 );
                                                                              								Aviso := False;
                                                      																			end;
                                                      																	end;
                                                      															end;
                                                      													end;
                                                                            end;
                                                      									end;
                                                      							end;
                                                      					end;
                                                      			end;
                                                      	end;
}
                                                      If (NOT ((DADORX[3] = 0) And (DADORX[4] = 0) And (DADORX[5] = 0))) then

                                                      	If (NOT ((DADORX[10] = 0) And (DADORX[11] = 0) And (DADORX[12] = 0))) then

                                                        	If (NOT ((DADORX[17] = 0) And (DADORX[18] = 0) And (DADORX[19] = 0))) then

                                                          	If (NOT ((DADORX[24] = 0) And (DADORX[25] = 0) And (DADORX[26] = 0))) then

                                                            	If (NOT ((DADORX[31] = 0) And (DADORX[32] = 0) And (DADORX[33] = 0))) then

                                                              	If (NOT ((DADORX[38] = 0) And (DADORX[39] = 0) And (DADORX[40] = 0))) then

                                                                	If (NOT ((DADORX[45] = 0) And (DADORX[46] = 0) And (DADORX[47] = 0))) then

                                                                  	If (NOT ((DADORX[52] = 0) And (DADORX[53] = 0) And (DADORX[54] = 0))) then

                                                                    	If (NOT ((DADORX[59] = 0) And (DADORX[60] = 0) And (DADORX[61] = 0))) then

                                                                      	If (NOT ((DADORX[66] = 0) And (DADORX[67] = 0) And (DADORX[68] = 0))) then
                                                                        	begin
                                                                          	TRANSMITE_BLOCO ( LISTA_FULL,0 );
                                                                            Aviso := False;
                                                                          end;

                     																	If (Aviso) then
                                                      	begin

																												  vRegistros := TStringList.Create;

                                                          AssignFile(S, CaminhoBatidas);
                                                          Rewrite(S);
                                                          vetor := 1;

                                                          While (TEXTO[vetor] <> '') Do
                                                          	begin
                                                          		Write(S, TEXTO[vetor]);
                                                              vLinha := Trim(TEXTO[vetor]);
                                                              vRegistro := Copy(vLinha, 31, 5);										// Número do Crachá
                                                              vRegistro := Copy(vLinha, 40, 1) + ',' + vRegistro;	// Entrada / Saída
                                                              vRegistro := vRegistro + ',' + Copy(vLinha, 1, 8);	// Data
                                                              vRegistro := vRegistro + ',' + Copy(vLinha, 10, 5);	// Hora
                                                             	vRegistro := vRegistro + ',' + Copy(vLinha, 37, 2);	// Nº relógio
                                                              vRegistros.Add(vRegistro);
                                                          		Vetor := (Vetor + 1);
                                                          	end;

                                                          WriteLn(S);
                                                          CloseFile(S);

																													vRegistros.SaveToFile('c:\registros.txt');						// Salvar Registros em arquivo.

																													vRegistros.SaveToFile( Gerar_Nome_Arquivo_Registros() );

                                                          Habilita_Clear	:= True;
                                                          BitBtn2.Enabled	:= True;
                                                          BitBtn1.Enabled	:= True;
                                                          Timer2.Enabled	:= True;
                                                          Timer3.Enabled	:= True;

                                                          Sleep(500);

                                                          MEMO1.Lines.Add(TimeToStr(TIME) + ' Registros Baixados!');

                                                          If (Application.MessageBox('Deseja realmente APAGAR todos os registros do ponto?', 'APAGAR DADOS', MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2)  = ID_YES) then
                                                          	If (Application.MessageBox('Tem certeza que deseja apagar?', 'APAGAR DADOS', MB_YESNO + MB_ICONWARNING + MB_DEFBUTTON2) = ID_YES) then
                                                          		TRANSMITE_BLOCO(APAGAR_BATIDAS, 0);


                                                        end;

                                                    end;

						APAGAR_BATIDAS_CONF 									: begin
                      																Application.MessageBox('Memória de registros apagada com sucesso!', 'Memória Apagada', MB_OK + MB_ICONINFORMATION);
                      																MEMO1.Lines.Add(TimeToStr(TIME) + ' Batidas apagadas!');
                      															end;
					end;																																																		// Final do Case

				end;																																																			// Final do Begin

		end;																																																					// Final do IF

end;
{-----------------------------------------------------------------------------------------------------------------------}


procedure TForm1.Timer2Timer(Sender: TObject);
begin
	TRANSMITE_BLOCO(PING, 0);
end;
{-----------------------------------------------------------------------------------------------------------------------}


procedure TForm1.Timer3Timer(Sender: TObject);
begin

	StatusBar1.Panels[0].Text := 'Status: Desconectado... Verifique as conexões!';

  If (NOT Conectado) then
  	begin
 			MEMO1.Lines.Add(TimeToStr(TIME) + ' Desconectado');
 			Conectado:=True;
		end;

	Transacao := False;

end;
{-----------------------------------------------------------------------------------------------------------------------}


procedure TForm1.BitBtn2Click(Sender: TObject);
begin

	If (Application.MessageBox('Deseja realmente APAGAR todos os registros do ponto?', 'APAGAR DADOS', MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2)  <> ID_YES) then Exit;

	If (Application.MessageBox('Tem certeza que deseja apagar?', 'APAGAR DADOS', MB_YESNO + MB_ICONWARNING + MB_DEFBUTTON2) <> ID_YES) then Exit;

  TRANSMITE_BLOCO(APAGAR_BATIDAS, 0);

end;
{-----------------------------------------------------------------------------------------------------------------------}


procedure TForm1.Timer4Timer(Sender: TObject);
begin
  if (CheckBox1.Checked) then
  begin
    if (SpinEdit1.Text = '') then
      SpinEdit1.Value := 60;

    If (QUANTIDADE_BATIDAS > 0) and (Conectado) then
    begin
      BitBtn1.Enabled	:= False;

      gauge2.Progress	:= 0;

      VETOR_FUNC := 1;

      Timer2.Enabled	:= False;
      Timer3.Enabled	:= False;
      
      TRANSMITE_BLOCO(LISTA_FULL_INI, 0);
    end;
  end;
end;

procedure TForm1.CheckBox1Click(Sender: TObject);
begin
  if (CheckBox1.Checked) then
  begin
    SalvaLinha(caminhoini, 'true', 4);
    Timer4.Enabled := true;
  end else
  begin
    SalvaLinha(caminhoini, 'false', 4);
    Timer4.Enabled := false;
  end;
end;

procedure TForm1.SpinEdit1Change(Sender: TObject);
begin
  try
    Timer4.Interval := SpinEdit1.Value * 1000;
  Except
    Timer4.Interval := 60000;
    SpinEdit1.Value := 60;
  end
end;

procedure TForm1.SpinEdit1Exit(Sender: TObject);
begin
  if (SpinEdit1.Text = '') then
    SpinEdit1.Value := 60;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  SalvaLinha(caminhoini, SpinEdit1.Text, 3);
  if (CheckBox1.Checked) then
    SalvaLinha(caminhoini, 'true', 4)
  else
    SalvaLinha(caminhoini, 'false', 4);  
end;

end.

{
MB_ICONEXCLAMATION,
MB_ICONWARNING	An exclamation-point icon appears in the message box.
MB_ICONINFORMATION, MB_ICONASTERISK	An icon consisting of a lowercase letter i in a circle appears in the message box.
MB_ICONQUESTION	A question-mark icon appears in the message box.
MB_ICONSTOP,MB_ICONERROR,MB_ICONHAND	A stop-sign icon appears in the message box.

  MB_APPLMODAL	The user must respond to the message box before continuing work in the window identified by the hWnd parameter. However, the user can move to the windows of other applications and work in those windows. Depending on the hierarchy of windows in the application, the user may be able to move to other windows within the application. All child windows of the parent of the message box are automatically disabled, but popup windows are not.MB_APPLMODAL is the default if neither MB_SYSTEMMODAL nor MB_TASKMODAL is specified.
MB_SYSTEMMODAL	Same as MB_APPLMODAL except that the message box has the WS_EX_TOPMOST style. Use system-modal message boxes to notify the user of serious, potentially damaging errors that require immediate attention (for example, running out of memory). This flag has no effect on the user's ability to interact with windows other than those associated with hWnd.
MB_TASKMODAL

MB_DEFAULT_DESKTOP_ONLY

The desktop currently receiving input must be a default desktop; otherwise, the function fails. A default desktop is one an application runs on after the user has logged on.

MB_HELP

Adds a Help button to the message box. Choosing the Help button or pressing F1 generates a Help event.

MB_RIGHT

The text is right-justified.

MB_RTLREADING

Displays message and caption text using right-to-left reading order on Hebrew and Arabic systems.

MB_SETFOREGROUND

The message box becomes the foreground window. Internally, Windows calls the SetForegroundWindow function for the message box.

MB_TOPMOST

The message box is created with the WS_EX_TOPMOST window style.

MB_SERVICE_NOTIFICATION         }
