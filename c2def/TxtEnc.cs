//TxtEncSample
//
//Copyright(C)2004 G-PROJECT http://www.gprj.net/

using System;
using System.Text;
using System.IO;

namespace G_PROJECT
{
	public class TxtEnc: IDisposable
	{
		
		//認識するコードの最大個数
		protected const int NumCode		=7;

		//各コードの番号定義
		public const int CD_JPN_ISO2022	=0;
		public const int CD_JPN_SJIS	=1;
		public const int CD_JPN_EUC		=2;
		public const int CD_UTF_8		=3;
		public const int CD_UTF_16		=4;
		public const int CD_UTF_32		=5;
		public const int CD_UTF_7		=6;

		//ExCode一覧
		public const int CDEX_FLG_ERROR						=1;//エラーなもので読み込むためにはFIX必要
		public const int CDEX_FLG_FIX						=2;//仕様上正しいが読み込むためにはFIX必要
		public const int CDEX_FLG_INFO						=4;//情報

		//UTF-7
		public const int CDEX_UTF_7_ERROR_BASE64				=CDEX_FLG_ERROR	+ 32;
		public const int CDEX_UTF_7_FIX_RFC2060				=CDEX_FLG_FIX	+ 64;
		//UTF-8
		public const int CDEX_UTF_8_ERROR_DOUBLE				=CDEX_FLG_ERROR	+ 32;
		public const int CDEX_UTF_8_ERROR_0XC0				=CDEX_FLG_ERROR	+ 64;
		public const int CDEX_UTF_8_INFO_BOM				=CDEX_FLG_INFO	+ 128;
		public const int CDEX_UTF_8_INFO_UCS4				=CDEX_FLG_INFO	+ 256;
		//UTF-16/32	
		public const int CDEX_UTF_INFO_BOM_BE				=CDEX_FLG_INFO	+ 32;
		public const int CDEX_UTF_INFO_BOM_LE				=CDEX_FLG_INFO	+ 64;
		//ISO2022JP
		public const int CDEX_ISO2022_INFO_ESC				=CDEX_FLG_INFO	+ 32;
		public const int CDEX_ISO2022_INFO_SISO				=CDEX_FLG_INFO	+ 64;
		public const int CDEX_ISO2022_INFO_JIS0208_1978		=CDEX_FLG_INFO	+ 128;
		public const int CDEX_ISO2022_INFO_JIS0208_1983		=CDEX_FLG_INFO	+ 256;
		public const int CDEX_ISO2022_INFO_JIS0208_1990_1	=CDEX_FLG_INFO	+ 512;
		public const int CDEX_ISO2022_INFO_JIS0213_2000_A1	=CDEX_FLG_INFO	+ 1024;
		public const int CDEX_ISO2022_INFO_JIS0213_2000_A2	=CDEX_FLG_INFO	+ 2048;
		public const int CDEX_ISO2022_INFO_JIS0201_1976_K	=CDEX_FLG_INFO	+ 4096;
		public const int CDEX_ISO2022_INFO_ASCII			=CDEX_FLG_INFO	+ 8192;
		public const int CDEX_ISO2022_INFO_JIS0201_1976_RS1	=CDEX_FLG_INFO	+ 16384;
		public const int CDEX_ISO2022_INFO_JIS0201_1976_RS2	=CDEX_FLG_INFO	+ 32768;
		public const int CDEX_ISO2022_INFO_JIS0208_1990_2	=CDEX_FLG_INFO	+ 65536;
		//EUC
		public const int CDEX_EUC_INFO_EXK					=CDEX_FLG_INFO	+ 32;

		protected string srcCodec;
		protected int srcCodecIndex;
		protected string srcText;
		protected byte[] srcByte;
		protected int[] srcCountCodec;
		protected int[] srcExCodec;
		protected int srcMaxRead;
		protected bool srcCodeBreak;

		public TxtEnc()
		{//コンストラクタ
			srcMaxRead=2048;
			srcCodeBreak=true;
		}

		public void Dispose()
		{
			// デストラクタみたいなもの
		
			GC.SuppressFinalize(this);
		}

		public string Text
		{
			get 
			{//何かしらの操作をしないと戻らない
				return srcText; 
			}
			set 
			{//Unicodeで突入
				//srcText=value;
				srcByte=Encoding.Unicode.GetBytes(value);
				srcCodec="utf-16";
			}
		}
		public int[] CountCodec
		{
			get
			{
				return srcCountCodec;
			}
		}
		public int[] CountExCodec
		{
			get
			{
				return srcExCodec;
			}
		}
		public int MaxRead
		{
			get 
			{
				return srcMaxRead; 
			}
			set 
			{
				srcMaxRead=value;
			}		
		}
		public int ExCodec
		{
			get 
			{
				return srcExCodec[srcCodecIndex]; 
			}
		}
		public bool CodeBreak
		{
			get 
			{
				return srcCodeBreak; 
			}
			set 
			{
				srcCodeBreak=value;
			}		
		}

		public string Codec
		{
			get 
			{
				return srcCodec; 
			}
			set 
			{//変換
				if(srcCodec!="")
				{
					double d;

					if (double.TryParse(srcCodec,System.Globalization.NumberStyles.Any,System.Globalization.NumberFormatInfo.InvariantInfo,out d))
					{
						if (double.TryParse(value,System.Globalization.NumberStyles.Any,System.Globalization.NumberFormatInfo.InvariantInfo,out d))
						{
							srcByte=Encoding.Convert(Encoding.GetEncoding(int.Parse(srcCodec)),Encoding.GetEncoding(int.Parse(value)),srcByte);
							srcCodec=value;
							srcText=Encoding.GetEncoding(int.Parse(srcCodec)).GetString(srcByte);
						}
						else
						{
							srcByte=Encoding.Convert(Encoding.GetEncoding(int.Parse(srcCodec)),Encoding.GetEncoding(value),srcByte);
							srcCodec=value;
							srcText=Encoding.GetEncoding(int.Parse(srcCodec)).GetString(srcByte);
						}
					}
					else
					{
						if (double.TryParse(value,System.Globalization.NumberStyles.Any,System.Globalization.NumberFormatInfo.InvariantInfo,out d))
						{
							srcByte=Encoding.Convert(Encoding.GetEncoding(srcCodec),Encoding.GetEncoding(int.Parse(value)),srcByte);
							srcCodec=value;
							srcText=Encoding.GetEncoding(srcCodec).GetString(srcByte);
						}
						else
						{
							srcByte=Encoding.Convert(Encoding.GetEncoding(srcCodec),Encoding.GetEncoding(value),srcByte);
							srcCodec=value;
							srcText=Encoding.GetEncoding(srcCodec).GetString(srcByte);
						}
					}
				}
			}
		}
		
		public void SaveToFile(string file,string codec)
		{//ファイルへ書き込み
			Codec=codec;
			SaveToFile(file);
		}
		public void SaveToFile(string file)
		{//ファイルへ書き込み
			
			FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write);
			BinaryWriter writer = new BinaryWriter(fs);
			writer.Write(srcByte);
			writer.Close();
		}


		public Encoding SetFromTextFile(string file)
		{//ファイルから読み込み
			long t=0;
			return SetFromTextFile(file,ref t);
		}

		public Encoding SetFromTextFile(string file,ref long filesize)
		{//ファイルから読み込み
			//---
			FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read);
			filesize=fs.Length;
			Encoding enc=SetFromStream(fs);
			fs.Close();
			return enc;
		}

		public Encoding SetFromStream(Stream s)
		{//ストリームから読み込み
			BinaryReader reader = new BinaryReader(s);
			srcByte= reader.ReadBytes((int)(s.Length));
			return SetFromByteArray(ref srcByte);
		}
	
		public Encoding SetFromByteArray(ref byte[]txt)
		{//判別
			int AS=	txt.Length-1,			//最大サイズ
				RP=	0,						//リードポインタ
				RS=	srcMaxRead;				//最大読み込みサイズ

			int[]	code=	new int[NumCode];		//Code
			int[]	excode=	new int[NumCode];		//Excode
			int		code2=	-1;						//確定コードindex(-1=確定してない 0～確定index)
			int		code3=	-1;						//コードの状態 -2=確定もしくは可能性 -1=疑惑なし 0～=疑惑index
			bool[]	code4=	new bool[NumCode];		//例外コード(True=ありえない False=通常)
			
			int		rcode;//初期化用
			srcByte=txt;




			//ReadSize調整
			if(AS==-1){return null;}
			else if(RS==0){RS=AS;}
			else if(RS>AS){RS=AS;}
		
			//判定BOMチェック 文字列の最初にBOMがある場合は確定
			//UTF-16/32
			if(Chk_UTF_16_32(ref txt,ref RP,RS,AS,ref code,ref excode,ref code2,ref code3,ref code4)==1)
			{

				while(RP<=RS)
				{//判定メインループ jis sjis euc utf7 utf8
					rcode=1024;
					if(txt[RP]==0x1B && code4[CD_JPN_ISO2022] == false)
					{//ISO2022
						//**DBG-OK
						rcode=Chk_JPN_ISO2022(ref txt,ref RP,RS,AS,ref code,ref excode,ref code2,ref code3,ref code4);
					}
					else if(txt[RP]>0x7F)
					{
						if(txt[RP]>0xBF && txt[RP]<0xFE && code4[CD_UTF_8]==false)//@@@@@ && code4[CD_UTF_8]==false追加
						{//UTF
							switch(Chk_UTF_8(ref txt,ref RP,RS,AS,ref code,ref excode,ref code2,ref code3,ref code4))
							{
								case 0:
									//ﾌﾞｯﾁｬｹ ｱﾘｴﾅｲ
									rcode=0;
									break;
								case 1:
									rcode=Chk_JPN_SJIS_EUC(ref txt,ref RP,RS,AS,ref code,ref excode,ref code2,ref code3,ref code4);
									break;
								case 2:
									rcode=2;
									break;
							}	

						}
						else
						{//SJIS EUC
							if(code4[CD_UTF_8]==false)
							{
								code4[CD_UTF_8]=true;//UTF-8ではない
								if(code[CD_UTF_8]!=0)
								{//ﾌﾞｯﾁｬｹ ｱﾘｴﾅｲ
									rcode=0;
								}
								else
								{
									rcode=Chk_JPN_SJIS_EUC(ref txt,ref RP,RS,AS,ref code,ref excode,ref code2,ref code3,ref code4);
								}
								
							}
							else
							{
								rcode=Chk_JPN_SJIS_EUC(ref txt,ref RP,RS,AS,ref code,ref excode,ref code2,ref code3,ref code4);
							}
						}

					}
					else
					{
						rcode=Chk_UTF_7(ref txt,ref RP,RS,AS,ref code,ref excode,ref code2,ref code3,ref code4);
					}
					switch(rcode)
					{
						case 0:
							//ReLoop:
								//ﾌﾞｯﾁｬｹ ｱﾘｴﾅｲ
								RP=0;
							code= new int[NumCode];
							excode= new int[NumCode];
							code2=-1;
							code3=-1;
							break;
						case 1:
							//ASCIIとかJISとか
							RP++;
							break;
						case 2:
							if(srcCodeBreak==true && code2!=-1){goto exitfor;}
							break;
						case 1024:
							//ASCIIとかJISとか
							RP++;
							break;
					}
				}
			}
			exitfor:
				if(code2==-1)
				{//確定でない
					if(code3>-1)
					{//疑惑確定
						code2=code3;
					}
					else//@@@@@@@@@@@@@@バグ修正
					{
						//code2=0;
						code2 = CD_JPN_SJIS;	//わからなかったらShift JISにするように修正
					}
					for(int i=1;i<NumCode;i++)
					{
						if(code[i] > code[code2]){code2=i;}
					}
				}
			srcExCodec=excode;
			srcCountCodec=code;
			srcCodecIndex=code2;
			switch(code2)
			{
				case CD_JPN_ISO2022:
					if((excode[CD_JPN_ISO2022] & CDEX_ISO2022_INFO_JIS0201_1976_K) ==CDEX_ISO2022_INFO_JIS0201_1976_K)
					{
						srcCodec="csISO2022JP";
					}
					else if((excode[CD_JPN_ISO2022] & CDEX_ISO2022_INFO_SISO) ==CDEX_ISO2022_INFO_SISO)
					{
						srcCodec="50222";
					}
					else
					{
						srcCodec="iso-2022-jp";
					}
					
					break;
				case CD_JPN_SJIS:
					srcCodec="shift_jis";
					break;
				case CD_JPN_EUC:
					if((excode[CD_JPN_EUC] & CDEX_EUC_INFO_EXK) ==CDEX_EUC_INFO_EXK)
					{
						srcCodec="20932";
					}
					else
					{
						srcCodec="euc-jp";
					}

					break;
				case CD_UTF_8:
					srcCodec="utf-8";
					break;
				case CD_UTF_7:
					srcCodec="utf-7";
					break;
				case CD_UTF_16:
					if((excode[CD_UTF_16] & CDEX_UTF_INFO_BOM_BE)==CDEX_UTF_INFO_BOM_BE)
					{
						srcCodec="unicodeFFFE";
					}
					else
					{
						srcCodec="utf-16";
					}
					break;
				case CD_UTF_32:
					if((excode[CD_UTF_32] & CDEX_UTF_INFO_BOM_BE)==CDEX_UTF_INFO_BOM_BE)
					{
						srcCodec="utf-32BE";
					}
					else
					{
						srcCodec="utf-32";
					}
				
					break;
				default:
					srcCodec="shift_jis";
					//srcCodec="iso-2022-jp";
					break;
			}
			try
			{
				//UTF-8 BOM対応
				if(code2== CD_UTF_8)
				{
					if((excode[CD_UTF_8] &CDEX_UTF_8_INFO_BOM)==CDEX_UTF_8_INFO_BOM)
					{
						return new UTF8Encoding(true);
					}
					else
					{
						return new UTF8Encoding(false);
					}
				}
				else
				{	
					double d;

					if (double.TryParse(srcCodec,System.Globalization.NumberStyles.Any,	System.Globalization.NumberFormatInfo.InvariantInfo,out d))					
					{
						return Encoding.GetEncoding(int.Parse(srcCodec));
					}
					else
					{
						return Encoding.GetEncoding(srcCodec);
					}
				}
			}
			catch
			{
				return null;
			}


		}

		

		public Encoding SetFromByteArray(ref byte[]txt,string sCodec)
		{//判別済み
			srcCodec=sCodec;
			srcByte=txt;
			return Encoding.GetEncoding(sCodec);
		}

		
		protected int Chk_UTF_7(ref byte[]txt,ref int RP,int RS,int AS,ref int[] code,ref int[] excode,ref int code2,ref int code3,ref bool[] code4)//@@@@add code4
		{//UTF-7判別
			/*
			UTF-7
			Index		CD_UTF_7	(6)

			Excodec		CDEX_UTF_7_ERRORBASE64=Base64Error(=が含まれている(UTF-7では省略されることになっている))
						CDEX_UTF_7_RFC2060=修正版UTF-7 RFC2060
			
			*/
			if(!(code2==-1 || code2==CD_UTF_7)){return 1;}//確定コード回避
			if(code4[CD_UTF_7]==true){return 1;}//除外コード回避
			int i;
			if(RP+2<=AS)
			{
				if(txt[RP]==0x2B && txt[RP+1]!=0x2D)//+-は除外
				{//+ BASE64突入
					for(i=RP+1;i<AS+1;i++)
					{
						if(txt[i]==0x2D ||	//-
							((txt[i]>=0x00 && txt[i]<0x20) || txt[i]==0x7F))//制御コード
						{//- BASE64復帰
							RP=i+1;
							code[CD_UTF_7]+=1;
							return 2;
						}
						if(txt[i]==0x3D)
						{//=判定@Base64規則より、エンコードすべきデータが2もしくは3バイトしかない場合=を付加する(UTF7では省略されることが決まってる)
							excode[CD_UTF_7]|=CDEX_UTF_7_ERROR_BASE64;
							if(i+1<=AS)
							{
								if(txt[i+1]==0x3D)
								{
									if(i+2<=AS)
									{
										if(txt[i+2]==0x2D ||	//-
											((txt[i+2]>=0x00 && txt[i+2]<0x20) || txt[i+2]==0x7F))//制御コード
										{
											RP=i+3;
											code[CD_UTF_7]+=1;
											return 2;
										}
									}

								}
								else if(txt[i+1]==0x2D ||	//-
									((txt[i+1]>=0x00 && txt[i+1]<0x20) || txt[i+1]==0x7F))//制御コード
								{
									RP=i+2;
									code[CD_UTF_7]+=1;
									return 2;
								}
							}
						}
						if(!(txt[i]==0x2B ||
							(txt[i]>0x2E && txt[i]<0x40) ||
							(txt[i]>0x40 && txt[i]<0x5B) ||
							(txt[i]>0x60 && txt[i]<0x7B)))
						{//BASE64範囲外
							code4[CD_UTF_7]=true;
							return 0;
						}
					}
					//BASE64から復帰しない
					code4[CD_UTF_7]=true;
					return 0;
				}
				if(txt[RP]==0x26) //修正UTF7用(RFC2060) & 
				{
					for(i=RP+1;i<AS+1;i++)
					{
						if(txt[i]==0x2D ||	//-
							((txt[i]>=0x00 && txt[i]<0x20) || txt[i]==0x7F))//制御コード
						{//- BASE64復帰
							RP=i+1;
							code[CD_UTF_7]+=1;
							excode[CD_UTF_7]|=CDEX_UTF_7_FIX_RFC2060;
							return 2;
						}
						if(txt[i]==0x3D)
						{//=判定@Base64規則より、エンコードすべきデータが2もしくは3バイトしかない場合=を付加する(UTF7では省略されることが決まってる)
							excode[CD_UTF_7]|=CDEX_UTF_7_ERROR_BASE64;
							if(i+1<=AS)
							{
								if(txt[i+1]==0x3D)
								{
									if(i+2<=AS)
									{
										if(txt[i+2]==0x2D ||	//-
											((txt[i+2]>=0x00 && txt[i+2]<0x20) || txt[i+2]==0x7F))//制御コード
										{
											RP=i+3;
											code[CD_UTF_7]+=1;
											excode[CD_UTF_7]|=CDEX_UTF_7_FIX_RFC2060;
											return 2;
										}
									}

								}
								else if(txt[i+1]==0x2D ||	//-
									((txt[i+1]>=0x00 && txt[i+1]<0x20) || txt[i+1]==0x7F))//制御コード
								{
									RP=i+2;
									code[CD_UTF_7]+=1;
									excode[CD_UTF_7]|=CDEX_UTF_7_FIX_RFC2060;
									return 2;
								}
							}
						}
						if(!(txt[i]==0x2B || txt[i]==0x2C ||
							(txt[i]>0x2F && txt[i]<0x40) ||
							(txt[i]>0x40 && txt[i]<0x5B) ||
							(txt[i]>0x60 && txt[i]<0x7B)))
						{//BASE64範囲外
							code4[CD_UTF_7]=true;
							return 0;
						}
					}
					//BASE64から復帰しない
					code4[CD_UTF_7]=true;
					return 0;

				}
				//
			}
			return 1;
			

			
			
		}

		protected int Chk_UTF_8(ref byte[]txt,ref int RP,int RS,int AS,ref int[] code,ref int[] excode,ref int code2,ref int code3,ref bool[] code4)
		{//UTF-8判別
			/*
			UTF-8
			Index		CD_UTF_8	(3)

			Excodec		CDEX_UTF_8_ERROR_DOUBLE	=ERROR		(Break UTF-8 double first byte)
						CDEX_UTF_8_INFO_BOM		=Include	(Byte Order Mark)
						CDEX_UTF_8_ERROR_0XC0	=Include	(0xC0 0xC1)
						CDEX_UTF_8_INFO_UCS4	=UCS-4		(RFC2279)
			*/
			if(!(code2==-1 || code2==CD_UTF_8)){return 1;}//確定コード回避
			if(code4[CD_UTF_8]==true){return 1;}//除外コード回避
			int size=0,
				i;
			bool brcode=false;
			if(txt[RP]>0xBF && txt[RP]<0xFE)
			{//UTF-8
			
				//可変長サイズ決定
				if((txt[RP] & 0xFC) == 0xFC)
				{
					size=5;
					excode[CD_UTF_8]|=CDEX_UTF_8_INFO_UCS4;//UCS-4
				}
				else if((txt[RP] & 0xF8) == 0xF8)
				{
					size=4;
					excode[CD_UTF_8]|=CDEX_UTF_8_INFO_UCS4;//UCS-4
				}
				else if((txt[RP] & 0xF0) == 0xF0){size=3;}
				else if((txt[RP] & 0xE0) == 0xE0)
				{
					size=2;
					if(RP+2<=AS)
					{
						if(txt[RP]==0xEF && txt[RP+1]==0xBB && txt[RP+2]==0xBF)
						{//BOM
							excode[CD_UTF_8]|=CDEX_UTF_8_INFO_BOM;
							code[CD_UTF_8]+=3;//BOM CNT
							RP+=3;
							return 2;
						}
					}

				}
				else if(txt[RP] == 0xC0 || txt[RP] == 0xC1)
				{
					//0xC0 0xC1 Security
					excode[CD_UTF_8]|=CDEX_UTF_8_ERROR_0XC0;
					size=1;
				}
				else if((txt[RP] & 0xC0) == 0xC0){size=1;}
			
				if(RP+size>AS)
				{//UTFで指定されてるサイズがオーバーしてる
					//UTF-8でない
					code4[CD_UTF_8]=true;
					return 0;
				}
				if(txt[RP+1]==txt[RP])
				{//ダブリバグ？（連続することがある
					size++;
					if(RP+size>AS)
					{//UTFで指定されてるサイズがオーバーしてる
						//utf-8でない
						code4[CD_UTF_8]=true;
						return 0;
					}
					brcode=true;
				}
				for(i=RP+1;i<=RP+size;i++)
				{
					if(!(txt[i] >0x7F && txt[i] <0xC0))
					{
						//範囲外のためUTF-8でない
						code4[CD_UTF_8]=true;
						return 0;
					}
				}
				if(brcode==true)
				{//壊れてるUTFの可能性ｱﾘ
					excode[CD_UTF_8]|=CDEX_UTF_8_ERROR_DOUBLE;
					code3=CD_UTF_8;
				}
				else
				{
					code[CD_UTF_8]++;//UTF-8判定
					code3=-2;
				}
				RP+=size+1;
				return 2;
			}
			return 1;
		}
		
		protected int Chk_UTF_16_32(ref byte[]txt,ref int RP,int RS,int AS,ref int[] code,ref int[] excode,ref int code2,ref int code3,ref bool[] code4)//@@@@add code4
		{//UTF-16/32判別
			/*
			UTF-16
			Index		CD_UTF_16	(4)

			Excodec		CDEX_UTF_BOM_BE=BE
						CDEX_UTF_BOM_LE=LE

			UTF-32
			Index		CD_UTF_32	(5)

			Excodec		CDEX_UTF_BOM_BE=BE
						CDEX_UTF_BOM_LE=LE
			*/

			//BOM判定のみ
			//BOMが出た場合確定
			//16/32のBOMはファイルの最初にある

			//if(!(code2==-1 || code2==CD_UTF_16)){return 1;}//確定コード回避
			//if(code4[CD_JPN_ISO2022]==true){return 1;}//除外コード回避
			if(RP+1<=AS)
			{
				if(txt[RP]==0x00 && txt[RP+1]==0x00)
				{//BE
					if(RP+3<=AS)
					{
						if(txt[RP+2]==0xFE && txt[RP+3]==0xFF)
						{//UTF-32
							code[CD_UTF_32]+=1;
							code2=CD_UTF_32;
							code3=-2;
							excode[CD_UTF_32]|=CDEX_UTF_INFO_BOM_BE;
							RP+=4;
							return 2;
						}
					}
				}
				if(txt[RP]==0xFE && txt[RP+1]==0xFF)
				{//BE
					//UTF-32CHK
					if(RP>1)
					{
						if(txt[RP-2]==0x00 && txt[RP-1]==0x00)
						{//UTF-32
							code[CD_UTF_32]+=1;
							code2=CD_UTF_32;
							code3=-2;
							excode[CD_UTF_32]|=CDEX_UTF_INFO_BOM_BE;
							RP+=2;
							return 2;
						}
					}

					code[CD_UTF_16]+=1;
					code2=CD_UTF_16;
					code3=-2;
					excode[CD_UTF_16]|=CDEX_UTF_INFO_BOM_BE;
					RP+=2;
					return 2;
				}
				if(txt[RP]==0xFF && txt[RP+1]==0xFE)
				{//LE
					//UTF-32CHK
					if(RP+3<=AS)
					{
						if(txt[RP+2]==0x00 && txt[RP+3]==0x00)
						{//UTF-32
							code[CD_UTF_32]+=1;
							code2=CD_UTF_32;
							code3=-2;
							excode[CD_UTF_32]|=CDEX_UTF_INFO_BOM_LE;
							RP+=4;
							return 2;
						}
					}
					code[CD_UTF_16]+=1;
					code2=CD_UTF_16;
					code3=-2;
					excode[CD_UTF_16]|=CDEX_UTF_INFO_BOM_LE;
					RP+=2;
					return 2;
				}
			}
			return 1;
			
		}
		protected int Chk_JPN_ISO2022(ref byte[]txt,ref int RP,int RS,int AS,ref int[] code,ref int[] excode,ref int code2,ref int code3,ref bool[] code4)
		{
			/*
			ISO-2022-JP
			Index		CD_JPN_ISO2022
			Excode		CDEX_ISO2022_ESC				=不明ESC
						CDEX_ISO2022_SISO				=SI/SO
						CDEX_ISO2022_JIS0208_1978		=JIS X 0208-1978
						CDEX_ISO2022_JIS0208_1983		=JIS X 0208-1983
						CDEX_ISO2022_JIS0208_1990_1		=JIS X 0208-1990
						CDEX_ISO2022_JIS0213_2000_A1	=JIS X 0213:2000 1面
						CDEX_ISO2022_JIS0213_2000_A2	=JIS X 0213:2000 2面
						CDEX_ISO2022_JIS0201_1976_K		=JIS X 0201-1976 片仮名
						CDEX_ISO2022_ASCII				=ASCII
						CDEX_ISO2022_JIS0201_1976_RS1	=JIS X 0201-1976 Roman Set(0x1B 0x28 0x4A)
						CDEX_ISO2022_JIS0201_1976_RS2	=JIS X 0201-1976 Roman Set(0x1B 0x28 0x48)
						CDEX_ISO2022_JIS0208_1990_2		=JIS X 0208-1990
						
			Code Area	ESC		=0x1B
						ESC中間	=0x20-0x2F
						ESC末端	=0x30-0x7E
			return
						0=ISO-2022でありえないコード出現
						1=ISO-2022じゃない
						2=ISO-2022ですよ
			*/ 

			if(!(code2==-1 || code2==CD_JPN_ISO2022)){return 1;}//確定コード回避
			
			if(code4[CD_JPN_ISO2022]==true){return 1;}//除外コード回避

			//SetByteArray以外から呼び出されることを考慮してRPを再チェック
			if(txt[RP]==0x0E ||txt[RP]==0x0F)
			{//SI SO判定
				excode[CD_JPN_ISO2022]|=CDEX_ISO2022_INFO_SISO;
			}
			else if(txt[RP]==0x1B)
			{//JIS エスケープシーケンス

				if(AS >=RP+2)
				{
					if (txt[RP+1]==0x24)
					{
						if(txt[RP+2]==0x40)
						{//0x1B 0x24 0x40 JIS X 0208-1978
							code[CD_JPN_ISO2022]+=1;
							excode[CD_JPN_ISO2022]|=CDEX_ISO2022_INFO_JIS0208_1978;
							code2=CD_JPN_ISO2022;
							code3=-2;
							RP+=3;
							return 2;
						}
						else if(txt[RP+2]==0x42)
						{//0x1B 0x24 0x42 JIS X 0208-1983
							code[CD_JPN_ISO2022]+=1;
							excode[CD_JPN_ISO2022]|=CDEX_ISO2022_INFO_JIS0208_1983;
							code2=CD_JPN_ISO2022;
							code3=-2;
							RP+=3;
							return 2;
						}
						else if(AS >=RP+3)
						{
							if(txt[RP+2]==0x28)
							{
								if(txt[RP+3]==0x44)
								{//0x1B 0x24 0x28 0x44 JIS X 0208-1990
									code[CD_JPN_ISO2022]+=1;
									excode[CD_JPN_ISO2022]|=CDEX_ISO2022_INFO_JIS0208_1990_1;
									code2=CD_JPN_ISO2022;
									code3=-2;
									RP+=4;
									return 2;
								}
								else if(txt[RP+3]==0x4F)
								{//0x1B 0x24 0x28 0x4F JIS X 0213:2000 1面
									code[CD_JPN_ISO2022]+=1;
									excode[CD_JPN_ISO2022]|=CDEX_ISO2022_INFO_JIS0213_2000_A1;
									code2=CD_JPN_ISO2022;
									code3=-2;
									RP+=4;
									return 2;
								}
								else if(txt[RP+3]==0x50)
								{//0x1B 0x24 0x28 0x50 JIS X 0213:2000 2面
									code[CD_JPN_ISO2022]+=1;
									excode[CD_JPN_ISO2022]|=CDEX_ISO2022_INFO_JIS0213_2000_A2;
									code2=CD_JPN_ISO2022;
									code3=-2;
									RP+=4;
									return 2;
								}
							}
						}
					}
					else if (txt[RP+1]==0x28)
					{
						if(txt[RP+2]==0x49)
						{//1B 28 49 JIS X 0201-1976 片仮名
							code[CD_JPN_ISO2022]+=1;
							excode[CD_JPN_ISO2022]|=CDEX_ISO2022_INFO_JIS0201_1976_K;
							code2=CD_JPN_ISO2022;
							code3=-2;
							RP+=3;
							return 2;
						}
						else if(txt[RP+2]==0x42)
						{//1B 28 42 ASCII
							code[CD_JPN_ISO2022]+=1;
							excode[CD_JPN_ISO2022]|=CDEX_ISO2022_INFO_ASCII;
							code2=CD_JPN_ISO2022;
							code3=-2;
							RP+=3;
							return 2;
						}
						else if(txt[RP+2]==0x4A)
						{//1B 28 4A JIS X 0201-1976 Roman Set
							code[CD_JPN_ISO2022]+=1;
							excode[CD_JPN_ISO2022]|=CDEX_ISO2022_INFO_JIS0201_1976_RS1;
							code2=CD_JPN_ISO2022;
							code3=-2;
							RP+=3;
							return 2;
						}
						else if(txt[RP+2]==0x48)
						{//1B 28 48 JIS X 0201-1976 Roman Set
							code[CD_JPN_ISO2022]+=1;
							excode[CD_JPN_ISO2022]|=CDEX_ISO2022_INFO_JIS0201_1976_RS2;
							code2=CD_JPN_ISO2022;
							code3=-2;
							RP+=3;
							return 2;
						}
					}
					else if(AS >=RP+5)
					{
						if(txt[RP+1]==0x26 &&
							txt[RP+2]==0x40 &&
							txt[RP+3]==0x1B &&
							txt[RP+4]==0x24 &&
							txt[RP+5]==0x42)
						{//0x1B 0x26 0x40 0x1B 0x24 0x42 JIS X 0208-1990
							code[CD_JPN_ISO2022]+=1;
							excode[CD_JPN_ISO2022]|=CDEX_ISO2022_INFO_JIS0208_1990_2;
							code2=CD_JPN_ISO2022;
							code3=-2;
							RP+=6;
							return 2;
						}
					}
				}

				//不明エスケープ
				//ESC=0x1B		ESC中間=0x20-0x2F		ESC末端=0x30-0x7E
				for(int i= RP+1;i<AS+1;i++)
				{
					if(txt[i] > 0x2F &&
						txt[i] < 0x7F)
					{//末端コード 発見(*ﾟДﾟ) ﾑﾎﾑﾎ 
						//不明ESCコード確定
						code[CD_JPN_ISO2022]+=1;
						excode[CD_JPN_ISO2022]|=CDEX_ISO2022_INFO_ESC;
						code2=CD_JPN_ISO2022;
						code3=-2;
						RP+=i-RP;
						return 2;
					}
					if(txt[i] < 0x20 ||
						txt[i] > 0x7E)
					{//ISO-2022じゃない
						code4[CD_JPN_ISO2022]=true;
						return 0;
					}
				}
				
			}
			return 1;
		}
		protected int Chk_JPN_SJIS_EUC(ref byte[]txt,ref int RP,int RS,int AS,ref int[] code,ref int[] excode,ref int code2,ref int code3,ref bool[] code4)
		{
			/*
			SJIS
			Index		CD_JPN_SJIS
			*/ 
			/*
			EUC
			Index		CD_JPN_EUC
			Excode		CDEX_EUC_EXK		=補助漢字

			*/ 
			if(!(code2==-1 || code2==CD_JPN_SJIS || code2==CD_JPN_EUC)){return 1;}//確定コード回避

			if(code4[CD_JPN_EUC]==true && code4[CD_JPN_SJIS]==true ){return 1;}//除外コード回避

			if(txt[RP]==0x8E && RP+1<=AS && (code2==CD_JPN_EUC || code2==-1) && code4[CD_JPN_EUC]==false)
			{//8EはEUCにおける半角のエスケープシークエンス
				if(txt[RP+1]>0xA0 && txt[RP+1]<0xE0 && 
					((code[CD_JPN_SJIS]==0 || code3!=-2)|| code3==CD_JPN_EUC)) //@@@変更　SJIS判定されてないことを追加
				{//a1->df範囲はsjisとの衝突
					code[CD_JPN_EUC]++;
					code3=-2;
					RP+=2;//とりあえずEUC判定して次へスルー(使われる範囲のSJISにおける漢字が単独ではあまり使われない)@@@@@@ RP++から+=2に変更
					return 2;

				}
				else if(txt[RP+1]>0x7F && txt[RP+1]<0xA1)
				{//SJIS確定
					code[CD_JPN_SJIS]++;
					code2=CD_JPN_SJIS;
					code3=-2;
					RP+=2;
					return 2;
				}
				else if ((txt[RP+1]>0x3F && txt[RP+1]<0x7F)||(txt[RP+1]>0x7F && txt[RP+1]<0xFD))
				{//もし第二バイトコードがSJIS範囲ならばSJIS　0x40<->0x7E 0x80<->0xFC @@@@@@@@@@@@@@@@@@@@@@
					code[CD_JPN_SJIS]++;
					code3=-2;
					RP+=2;
					return 2;
				}
				else
				{//不明コード
					//EUCではない
					code4[CD_JPN_EUC]=true;
					return 0;

				}
			}
			
			//補助漢字が入らない場合はここからコメントアウト
			if(txt[RP]==0x8F && RP+2<=AS && (code2==CD_JPN_EUC || code2==-1) && code4[CD_JPN_EUC]==false)
			{//8EはEUCにおける補助漢字のエスケープシークエンス
				if(txt[RP+1]>0xA0 && txt[RP+1]<0xFF && 
					txt[RP+2]>0xA0 && txt[RP+2]<0xFF && 
					((code[CD_JPN_SJIS]==0 || code3!=-2)|| code3==CD_JPN_EUC))
				{

					code[CD_JPN_EUC]++;
					code3=-2;
					excode[CD_JPN_EUC]|=CDEX_EUC_INFO_EXK;

					if((txt[RP+1]==0xFD || txt[RP+1]==0xFE)||
						(txt[RP+2]==0xFD || txt[RP+2]==0xFE))
					{//0xFD 0xFE EUC美乳
						RP+=3;						
						code2=CD_JPN_EUC;
						return 2;//breakと同意
					}


					RP+=3;//とりあえずEUC判定して次へスルー(とりあえず他の可能性もあるので+1)@@@@@@ RP++から+=3に変更
					return 2;
						
				}
				else if(txt[RP+1]>0x7F && txt[RP+1]<0xA1)
				{//SJIS確定
					code[CD_JPN_SJIS]++;
					code2=CD_JPN_SJIS;
					code3=-2;
					RP+=2;
					return 2;
				}
				else if ((txt[RP+1]>0x3F && txt[RP+1]<0x7F)||(txt[RP+1]>0x7F && txt[RP+1]<0xFD))
				{//もし第二バイトコードがSJIS範囲ならばSJIS　0x40<->0x7E 0x80<->0xFC @@@@@@@@@@@@@@@@@@@@@@
					code[CD_JPN_SJIS]++;
					code3=-2;
					RP+=2;
					return 2;
				}
				else
				{//不明コード
					//EUCではない
					code4[CD_JPN_EUC]=true;
					return 0;
				}
			}
				//補助漢字が入らない場合はここまでコメントアウト
			
			else
			{

				if(txt[RP]>0x7F && txt[RP]<0xA1 && (code2==CD_JPN_SJIS || code2==-1) && code4[CD_JPN_SJIS]==false)
				{//SJIS確定
					code[CD_JPN_SJIS]++;
					code2=CD_JPN_SJIS;
					code3=-2;
					RP++;
					return 2;
				}

				else if(txt[RP]>0xA0 && txt[RP]<0xE0 && 
					(code2==-1 || code2==CD_JPN_SJIS) && 
					((code[CD_JPN_EUC]==0 || code3!=-2) || code3==CD_JPN_SJIS) && //@@@変更　EUC判定されてないことを追加
					code4[CD_JPN_SJIS]==false)
				{//SJIS半角
					if(RP+1<=AS)
					{
						if((txt[RP]==0xA4 && (txt[RP+1]>0xA0 && txt[RP+1]<0xF4))
							||(txt[RP]==0xA5 && (txt[RP+1]>0xA0 && txt[RP+1]<0xF7)) && code4[CD_JPN_EUC]==false)
						{
							//EUCの場合はA4+?でかな A5+?でカナになる SJIS半角カナ条件にはいる
							//この条件のみSJIS半角から除外する
							//EUC全角カナ？
							code[CD_JPN_EUC]+=2;
							code3=CD_JPN_EUC;
							RP+=2;
							return 2;

						}
						else if(txt[RP+1]>0xDF && txt[RP+1]<0xFF && code4[CD_JPN_EUC]==false)
						{//SJIS半角カナ以外でのEUC範囲がでたためEUC
							code[CD_JPN_EUC]++;
							code3=-2;
							if(RP+2<=AS)
							{
								if((txt[RP+1]==0xFD || txt[RP+1]==0xFE)||
									(txt[RP+2]==0xFD || txt[RP+2]==0xFE))
								{//0xFD 0xFE EUC美乳
									RP+=2;
									code2=CD_JPN_EUC;
									return 2;//breakと同意
								}
							}

							RP++;//とりあえずEUC判定して次へスルー(とりあえず他の可能性もあるので+1)
							return 2;
						}
						else
						{//EUC外
							//SJIS半角の可能性
							code[CD_JPN_SJIS]++;
							code3=-2;
							RP++;
							return 2;
						}
					}
					else
					{
						//SJISの可能性
						code[CD_JPN_SJIS]++;
						code3=-2;
						RP++;
						return 2;
					}
				}
				else if(txt[RP]>0xA0 && txt[RP]<0xFF && (code2==CD_JPN_EUC || code2==-1) && code4[CD_JPN_EUC]==false)
				{//EUC
					code[CD_JPN_EUC]++;
					code3=-2;
					if(txt[RP]==0xFD || txt[RP]==0xFE)
					{//0xFD 0xFE EUC
						code2=CD_JPN_EUC;
					}
					RP++;
					return 2;
				}

				else
				{
					return 1;
				}

			}						
			//			return false;
		}
	}
}
