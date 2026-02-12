#include"EMReadWriteCard.h"
#include"Uart.h"

///////////////////////////////////////////////////////////////////
//                        Ð´¿¨Ê±¼ä²ÎÊýus
#define  	Start_Gap	 	  55		//¿ªÊ¼¼äÏ¶	
#define 	Write_Data1		32		//Ð´Êý¾Ý'1'
#define 	Write_DataH0	17		//Ð´Êý¾Ý'0'17 
#define 	Write_DataL0	23		//Ð´Êý¾Ý'0'23

unsigned char EM_Buff[10];		//¶Á¿¨Êý¾Ý½ÓÊÕ»º³å
unsigned int Card_Periodtime = 0;
unsigned int Card_JumpTime = 0;
	
void Delay_8nus(unsigned int n)
{
		unsigned char i;
	
		do
		{
				i = 58;
				while (--i);
		}
		while(n--);	 
}

void Delay_nus(unsigned int n)
{
		TL0 = 0;TH0 = 0;TR0 = 1;
		while((TH0 * 256 + TL0) < n);
	  TR0 = 0;
}

unsigned int LineCheck(unsigned char u_Byte, unsigned char u_Len)
{
	  unsigned char x;
	  unsigned char Check_temp = 0;	
	  unsigned int Check = 0;
	
	  Check = (unsigned int)u_Byte & 0x00FF;
	
		for(x = 0; x < u_Len; x ++)
	  {
			  if(u_Byte & 0x01)		Check_temp ++;
			  u_Byte = u_Byte >> 1;
		}
		
		if((Check_temp % 2) != 0)
    {
				switch(u_Len)
				{
						case 0: 
								break;
						case 1: Check = Check | 0x0002;
								break;
						case 2: Check = Check | 0x0004;										 
								break;
						case 3: Check = Check | 0x0008;
								break;
						case 4: Check = Check | 0x0010;
								break;
						case 5: Check = Check | 0x0020;
								break;
						case 6: Check = Check | 0x0040;
								break;	
						case 7: Check = Check | 0x0080;						
								break;	
						case 8: Check = Check | 0x0100;						
								break;							
				}			
		}			
		return Check; 
}

unsigned char Columncheck(unsigned char *b_Byte, unsigned char Byte_Len)		//ÁÐÐ£Ñé
{
	  unsigned char x,y;
	  unsigned char Check = 0;	
	  unsigned char Check_temp = 0;
	  
		for(x = 0;x < 8;x ++)
		{
				Check = Check << 1;
			  Check_temp = 0;
			
				for(y = 0;y < Byte_Len;y ++)
				{
						switch(x)
						{
								case 0: Check_temp = Check_temp ^ (b_Byte[y] & 0x80);
										break;
								case 1: Check_temp = Check_temp ^ (b_Byte[y] & 0x40);
										break;
								case 2: Check_temp = Check_temp ^ (b_Byte[y] & 0x20);
										break;
								case 3: Check_temp = Check_temp ^ (b_Byte[y] & 0x10);
										break;
								case 4: Check_temp = Check_temp ^ (b_Byte[y] & 0x08);
										break;			
								case 5: Check_temp = Check_temp ^ (b_Byte[y] & 0x04);
										break;
								case 6: Check_temp = Check_temp ^ (b_Byte[y] & 0x02);
										break;
								case 7: Check_temp = Check_temp ^ (b_Byte[y] & 0x01);
										break;								
						}
				}
				if(Check_temp != 0) Check = Check | 0x01;
		}
		return Check; 
}
//========================================================================
// º¯Êý: void Write_bit(unsigned char x)
// ÃèÊö: Ð´Ò»¸öÊý¾ÝÎ»
// ²ÎÊý: x£ºÐ´ÈëµÄÊý¾Ý 0/1
// ·µ»Ø: 
// °æ±¾: V1.0, 2017-9-23
//========================================================================
void Write_bit(unsigned char x)
{
    if(x)
    {
				Carrier_On(); 						//¿ª´Å³¡
				Delay_8nus(Write_Data1);	//Î¬³ÖÊý¾Ý1µÄ´Å³¡Ê±¼ä
    }
    else
    {
				Carrier_On(); 						//¿ª´Å³¡
				Delay_8nus(Write_DataH0);	//Î¬³ÖÊý¾Ý0µÄ¿ªÆô´Å³¡Ê±¼ä
				Carrier_Off();						//¹Ø±Õ´Å³¡
				Delay_8nus(Write_DataL0);	//Î¬³ÖÊý¾Ý0µÄ¹Ø±Õ´Å³¡Ê±¼ä 	
        Carrier_On(); 						//¿ª´Å³¡    	
    } 
}

//========================================================================
// º¯Êý: void Write_star(void)
// ÃèÊö: ·¢ÆðÊ¼ÐÅºÅ
// ²ÎÊý: 
// ·µ»Ø: 
// °æ±¾: V1.0, 2017-9-23
//========================================================================
void Write_star(void)
{	
		Carrier_Off();
		Delay_8nus(Start_Gap); 			//¿ªÆô³¡Ç¿410us
		Write_bit(0);
}

//========================================================================
// º¯Êý: void Write_command(unsigned char command)
// ÃèÊö: Ð´ÃüÁî×Ö
// ²ÎÊý: command  Ð´£º0101 ¶Á£º1001 µÇÂ¼£º0011 ±£»¤£º1100 ½ûÓÃ£º1010º
// ·µ»Ø: 
// °æ±¾: V1.0, 2017-9-23
//========================================================================
void Write_command(unsigned char command)
{
		unsigned char i;
		unsigned char mask = 0x08;

		for(i=0;i<4;i++)								
		{
				Write_bit(mask & command);				 
				mask >>= 1;
		}	
}
//========================================================================
// º¯Êý: void WriteNbtye(unsigned int dat,unsigned char bit_len)
// ÃèÊö: Ð´Êý¾Ý
// ²ÎÊý: dat Ð´ÈëµÄÖµ
// ·µ»Ø: 
// °æ±¾: V1.0, 2017-9-23
//========================================================================
void WriteNbtye(unsigned int dat,unsigned char bit_len)
{
		unsigned char i;	
	
	  for(i = 0;i < bit_len;i ++)       //1¸öÐ£ÑéÎ»  ×î¸ßÎ»
	  {
				Write_bit(dat & 0x0001);
			  dat = dat >> 1;
		}
}

unsigned char EM43xxPatt(void)
{
	  unsigned char Patt_dat = 0;
	  unsigned char scan_temp = 250;
	  unsigned char ErrTimes = 0;
	  unsigned int  Old_Time = 0;
		unsigned int 	Period_Max = 0;
	
	  Period_Max = Card_Periodtime + (Card_Periodtime / 2);
		bitin = INPORT;
	
	  while(--scan_temp)
		{
				TL0 = 0;TH0 = 0;TR0 = 1;		//³õÊ¼»¯¶¨Ê±¼ÆÊýÆ÷ ¿ªÊ¼¼ÆÊ±	
				ErrTimes = 0;
				while(INPORT == bitin)
				{
						ErrTimes ++;
						Delay_8nus(1);
					  if(ErrTimes == 125) break; //1ms
				}
				bitin = INPORT;
				if(ErrTimes == 125)
				{
						Patt_dat = 0;
						continue; 
				}
				Old_Time = TH0 * 256 + TL0;
				TL0 = 0;TH0 = 0;
				if((Old_Time > (Card_Periodtime / 2)) && (Old_Time < Period_Max)) 		 
				{
						Patt_dat ++;
						if(Patt_dat == 4)
						{
								TL0 = 0;TH0 = 0;TR0 = 0;
								return 1;
						}
				}
				else
				{
						Patt_dat = 0;
				}
		}
		TL0 = 0;TH0 = 0;TR0 = 0;
		
		return 0;	
}

bit CheckDataEM4305(void)
{
		unsigned char i,j;
		unsigned char idata B_buff[45] = {0};
		unsigned int idata Check_B[4] = {0};
	
		if(!EM43xxPatt())   return 0;
	  Delay_nus(Card_JumpTime);      //ÑÓÊ±3/4ÖÜÆÚÊ±¼ä (nRF * 3) / 4  * 8;
		
		for(i = 0 ; i < 45; i ++)
		{
				bitin = INPORT;
				if(bitin)  B_buff[i] = 1; 
			  while(INPORT == bitin);		//µÈ´ýÌø±ä
			  Delay_nus(Card_JumpTime); //ÑÓÊ±3/4ÖÜÆÚÊ±¼ä (nRF * 3) / 4  * 8;
		}
		
    for(i = 0; i < 4; i ++)         //4¸ö×Ö½ÚÊý¾Ý    
    {
         EM_Buff[3 - i] = 0;
         for(j = 0; j < 8; j ++)    //ÍùÒ»¸ö×Ö½ÚÐ´Î»
         {
              EM_Buff[3 - i] = EM_Buff[3 - i] >> 1;
              if(B_buff[i * 9 + j])  EM_Buff[3 - i] |= 0x80;
         }
    }
		EM_Buff[5] = 0;
    for(i = 0; i < 4; i ++)
    {
        EM_Buff[5] = EM_Buff[5] << 1;			//±£´æÐÐÐ£Ñé
        if(B_buff[i * 9 + 8])  EM_Buff[5] |= 0x01;
    }		
		EM_Buff[4] = 0;
		for(i = 0; i < 8; i ++)		//±£´æÁÐÐ£Ñé
    {
        EM_Buff[4] = EM_Buff[4] >> 1;
        if(B_buff[36 + i])  EM_Buff[4] |= 0x80;     
    }
			
		for(i = 0; i < 4; i ++)  Check_B[i] = LineCheck(EM_Buff[i], 8);  //¼ÆËãµÃµ½ÐÐÐ£ÑéÎ»ÔÚµÚ8bit
		for(i = 0; i < 4; i ++)
		{
			  Check_B[i] = (Check_B[i] >> 8) & 0x00FF;
			  B_buff[0] = (EM_Buff[5] >> i) & 0x01;
			  if(Check_B[i] != B_buff[0])  return 0;
		}
		B_buff[0] = Columncheck(EM_Buff, 4);

		if(B_buff[0] != EM_Buff[4])			return 0;
		return 1;
}

void Set_ReadOperation(unsigned char _nRF)
{
		switch(_nRF)
		{
				case 8:	Card_Periodtime = (64 * (FOSC / 1000)) / 12 /1000;
								Card_JumpTime = (48 * (FOSC / 1000)) / 12 /1000;
								break;
				case 16:Card_Periodtime = (128 * (FOSC / 1000)) / 12 /1000;
								Card_JumpTime = (96 * (FOSC / 1000)) / 12 /1000;
								break;
				case 32:Card_Periodtime = (256 * (FOSC / 1000)) / 12 /1000;
								Card_JumpTime = (192 * (FOSC / 1000)) / 12 /1000;
								break;
				case 40:Card_Periodtime = (320 * (FOSC / 1000)) / 12 /1000;
								Card_JumpTime = (240 * (FOSC / 1000)) / 12 /1000;
								break;	
				case 64:Card_Periodtime = (512 * (FOSC / 1000)) / 12 /1000;
								Card_JumpTime = (384 * (FOSC / 1000)) / 12 /1000;
								break;	
				default:Card_Periodtime = (512 * (FOSC / 1000)) / 12 /1000;
								Card_JumpTime = (384 * (FOSC / 1000)) / 12 /1000;
								break;	
		}
}

unsigned char ReadWriteEM4305(unsigned char _cardCom,unsigned char _block,unsigned long _blockData,unsigned char _nRF)
{
		unsigned char i;
	  unsigned int  blockBuff[5] = {0};
		union 
		{
				unsigned long L;					
				unsigned char B[4];
		}_enCard;	
		_enCard.L = _blockData;
		
		for(i = 0;i < 4;i ++) blockBuff[i] = LineCheck(_enCard.B[i], 8);	//¼ÆËãÐÐÐ£Ñé	
		blockBuff[4] = Columncheck(_enCard.B,4); 	  //¼ÆËãÁÐÐ£Ñé
		_block = LineCheck(_block & 0x0F, 6);	        //¼ÆËãµØÖ·ÐÐÐ£Ñé 0100 0111
		
		Set_ReadOperation(_nRF);
		
		Write_star();		 	    			//Ð´ÆðÊ¼ÐÅºÅ
		switch(_cardCom)
		{
				case LoginCom:Write_command(LoginCom);         //Ð´µÇÂ½ÃüÁî ÑéÖ¤ÃÜÂë
											break;
				case ReadCom:	Write_command(ReadCom);					 //Ð´¶ÁÃüÁî	
											WriteNbtye(_block,7);      			 //·¢ËÍÐ´ÈëµÄµØÖ·  6bit + 1bitÐ£Ñé		
											return CheckDataEM4305();        //¶ÁÊý¾Ý·µ»Ø
        case WriteCom:Write_command(WriteCom);         //·¢ËÍÐ´ÃüÁî ¿éÐ´
                      WriteNbtye(_block,7);            //·¢ËÍÐ´ÈëµÄµØÖ·  6bit + 1bitÐ£Ñé
                      break;   	
			  case ProtectCom: 
                      Write_command(ProtectCom);       //Ð´±£»¤ÃüÁî Ð´OTP¿é
                      break;
        case DisableCom:
                      Write_command(DisableCom);       //Ð´½ûÓÃÃüÁî
                      for(i = 0;i < 4;i ++)  blockBuff[i] = 0x00FF;
					            blockBuff[4] = 0;
                      break;					
		}
		for(i = 0;i < 4;i ++)	WriteNbtye(blockBuff[i],9);    //Ð´Êý¾Ý 8bit + 1bitÐ£Ñé	
		WriteNbtye(blockBuff[4],8); //Ð´ÁÐÐ£Ñé 8bit
		Write_bit(0);         	     //Ð´½áÊø·û¡°0¡±
		if(_cardCom == DisableCom)   return 1;
		if(EM43xxPatt())  return 1;
		return 0;     //²Ù×÷Ê§°Ü
}















