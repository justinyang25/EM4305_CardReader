#include"STC8F.h"
#include"IntPut_IO.h"
#include"Uart.h"
#include"EM41xx.h"
#include"EMReadWriteCard.h"
#include"CopyEM4100.h"
#include"CRC16.h"

bit LED_flag = 0;
unsigned int  Time_Up = 0;
unsigned int  EM_Wait = 0;

void Delay125us(void)			//@24.000MHz
{
		unsigned char i, j;
	
		_nop_();
		_nop_();
		i = 4;
		j = 128;
		do
		{
				while (--j);
		}while (--i);
}
void Delay100ms(void)		 //@22.1184MHz
{
		unsigned char i, j, k;

		i = 12;
		j = 57;
		k = 122;
		do
		{
				do
				{
						while (--k);
				}while (--j);
		}while (--i);
}
//========================================================================
// º¯Êý: void BP_BP(unsigned char Counter,bit Mode)	
// ÃèÊö: ·äÃùÆ÷º¯Êý
// ²ÎÊý: Counter£º ´ÎÊý   Mode£º1³¤Ïì 0¶ÌÏìº
// ·µ»Ø: 
// °æ±¾: V1.0, 2017-9-23
//========================================================================
void BP_BP(unsigned char Counter,bit Mode)	
{
		unsigned char x;
	  unsigned int y,CheckCount = 0;
	
	  CheckCount = 200;
	  if(Mode)  CheckCount = 1000;
	
    EA = 0;
		for(x = 0; x < Counter; x++)
		{
				for(y = 0;y < CheckCount;y ++)
				{
						BP = !BP;
						Delay125us();
				}
				BP = 0;
				if(x == (Counter - 1)) 	break;
				Delay100ms();
		}
		EA = 1;	
}

void Timer3InIt(void)		 //¶¨Ê±1ms
{  
    T3L = 0x30;          //65536-11.0592M/12/1000
    T3H = 0xF8;
    T4T3M = 0x08;        //Æô¶¯¶¨Ê±Æ÷ 12TÄ£Ê½
    IE2 = ET3;           //Ê¹ÄÜ¶¨Ê±Æ÷ÖÐ¶Ï
}
void SYS_InIt(void)
{
		BP = 0;
		LED = 0;
	  LED1 = 1;
	  T1CLK = 0;
	  INPORT = 1;
	
		P1M0 = 0x4C; 			     //0100 1100
		P1M1 = 0x00; 	
		P3M0 = 0x12; 			     //0001 0010
		P3M1 = 0x00; 	         //0000 0000
	
		UartInit();	
	  Timer3InIt();
	  T1Clk_Out(125);				 //Ê±ÖÓÊä³ö125K
	  RFID_Timer0InIt();     //¶¨Ê±Æ÷0³õÊ¼»¯£¬ÓÃÓÚRFID½ÅµçÆ½¼ÆÊý
	  BP_BP(1,0);
}

void main(void)
{
		union 
		{
				unsigned int  _uint;					
				unsigned char B[2];
		}xdata CRC16_Base;	
		union 
		{
				unsigned int  _uint;					
				unsigned char B[2];
		}xdata Block_Base;
		union 
		{
				unsigned long uL;					
				unsigned char B[4];
		}xdata BankData;	
		
		unsigned char i,j;
	  unsigned char Uart_Len;
	  unsigned char xdata Uart_Buff[40]= {0};
		unsigned char xdata ListEM4100UID[5] = {0};
		unsigned char xdata Read_PLL = 0;
		unsigned char xdata WriteCount = 0;
			
		SYS_InIt();
	
	  while(1)
		{
				if(cmdArrived)	
				{
						cmdArrived = 0;
						Uart_Len = UartRead(Uart_Buff, sizeof(Uart_Buff)); 	//½«½ÓÊÕµ½Êý¾Ý¶ÁÈ¡µ½»º´æ
						if((Uart_Buff[0] == 0xAA) && (Uart_Buff[Uart_Len - 1] == 0xEE))	//ºÏ¸ñµÄÊý¾ÝÍ·ºÍÎ²
						{
								CRC16_Base._uint = GetCRC16(Uart_Buff, Uart_Len - 3);  
								if((Uart_Buff[Uart_Len - 2] == CRC16_Base.B[0]) && (Uart_Buff[Uart_Len - 3] == CRC16_Base.B[1])) //CRC16Ð£ÑéÕýÈ·
								{	
										LED = 1;
										if((Uart_Buff[1] == 0xFF) && (Uart_Buff[2] == 0xFF))	 //Áª»ú
										{
												LED_flag = 1;
												BP_BP(1,0);
										}
										else
										{
												switch(Uart_Buff[1])
												{
														case 0xF0:Read_PLL = Uart_Buff[2];
																			for(i = 0; i < 2; i ++) Block_Base.B[i] = Uart_Buff[3 + i];	//µØÖ·
																			for(i = 0; i < 16; i ++) 
																			{
																					if(Block_Base._uint & 0x0001)    //¼ì²â¶ÁµÄµØÖ·
																					{
																							Uart_Buff[2] = ReadWriteEM4305(ReadCom,i,0,Read_PLL);
																						  Uart_Buff[3] = i;
																							Uart_Len = 4;	
																							if(Uart_Buff[2])
																							{
																									for(j = 0; j < 4; j ++)  Uart_Buff[Uart_Len ++] = EM_Buff[j];
																									BP_BP(1,0);
																							}
																							else
																							{
																									for(j = 0; j < 4; j ++)  Uart_Buff[Uart_Len ++] = 0;
																									BP_BP(2,0);																								
																							}
																							CRC16_Base._uint = GetCRC16(Uart_Buff,Uart_Len); 	
																							Uart_Buff[Uart_Len ++] = CRC16_Base.B[0];
																							Uart_Buff[Uart_Len ++] = CRC16_Base.B[1];
																							Uart_Buff[Uart_Len ++] = 0xEE;
																							UartWrite(Uart_Buff,Uart_Len);	 //·µ»ØÊý¾Ý		
																							EM_Wait = 100;
																							while(EM_Wait > 0);																							
																					}
																					Block_Base._uint = Block_Base._uint >> 1;
																			}
																			break;
														case 0xF1:for(i = 0; i < 2; i ++) Block_Base.B[i] = Uart_Buff[2 + i];	//µØÖ·
																			WriteCount = 0;
																			for(i = 0; i < 16; i ++)
																			{
																					if(Block_Base._uint & 0x0001)    //¼ì²âÐ´µÄµØÖ·
																					{
																							for(j = 0; j < 4; j ++)	BankData.B[3 - j] = Uart_Buff[4 + (WriteCount * 4) + j];
																							WriteCount ++;
																							Uart_Buff[2] = ReadWriteEM4305(WriteCom,i,BankData.uL,Read_PLL);
																							Uart_Buff[3] = i;  	//·µ»ØµØÖ·
																							Uart_Len = 4;
																							if(Uart_Buff[2])
																							{
																									BP_BP(1,0);
																							}
																							else
																							{
																									BP_BP(2,0);																								
																							}
																							CRC16_Base._uint = GetCRC16(Uart_Buff,Uart_Len); 	
																							Uart_Buff[Uart_Len ++] = CRC16_Base.B[0];
																							Uart_Buff[Uart_Len ++] = CRC16_Base.B[1];
																							Uart_Buff[Uart_Len ++] = 0xEE;
																							UartWrite(Uart_Buff,Uart_Len);	 //·µ»ØÊý¾Ý		
																					}
																					Block_Base._uint = Block_Base._uint >> 1;
																			}
																			break;
														case 0xF2:				//ÑéÖ¤ÃÜÂë µÇÂ¼ÃüÁî
																			for(j = 0; j < 4; j ++)	BankData.B[3 - j] = Uart_Buff[4 + j];
																			Uart_Buff[2] = ReadWriteEM4305(LoginCom,4,BankData.uL,Read_PLL);
																			Uart_Len = 3;
																			if(Uart_Buff[2])
																			{
																					BP_BP(1,0);
																			}
																			else
																			{
																					BP_BP(2,0);																								
																			}
																			CRC16_Base._uint = GetCRC16(Uart_Buff,Uart_Len); 	
																			Uart_Buff[Uart_Len ++] = CRC16_Base.B[0];
																			Uart_Buff[Uart_Len ++] = CRC16_Base.B[1];
																			Uart_Buff[Uart_Len ++] = 0xEE;
																			UartWrite(Uart_Buff,Uart_Len);	 //·µ»ØÊý¾Ý															
																			break;
														case 0xF3:Carrier_On();	 				//´ò¿ª´Å³¡	¶ÁEM4100¿¨
																			EM_Wait = 1000;
																			Uart_Buff[1] = 0xF3;
																			Uart_Buff[2] = 0;
																			Uart_Len = 3;	
																			while(EM_Wait > 0)
																			{
																					if(ReadEM41xxCardNo())
																					{
																							Uart_Buff[2] = 1;
																							BP_BP(1,0);
																							for(i = 0; i < 5; i ++) Uart_Buff[Uart_Len ++] = ID[i];
																							break;
																					}
																			}
																			if(Uart_Buff[2] == 0)
																			{
																					BP_BP(2,0);
																					for(i = 0; i < 5; i ++) Uart_Buff[Uart_Len ++] = 0;																			
																			}
																			CRC16_Base._uint = GetCRC16(Uart_Buff,Uart_Len); 	
																			Uart_Buff[Uart_Len ++] = CRC16_Base.B[0];
																			Uart_Buff[Uart_Len ++] = CRC16_Base.B[1];
																			Uart_Buff[Uart_Len ++] = 0xEE;
																			UartWrite(Uart_Buff,Uart_Len);	 //·µ»ØÊý¾Ý																					
																			break;
														case 0xF4:				//¸´ÖÆEM4100							
																			for(i = 0; i < 5; i ++)
																			{
																					Buff[i] = Uart_Buff[3 + i];
																					ListEM4100UID[i] = Buff[i];
																			}
																			switch(Uart_Buff[2])
																			{
																					case 0:         //»ù¿¨ÊÇ57¿¨
																									T5557CopyEM4100(Buff);
																								break;
																					case 1:         //»ù¿¨ÊÇ43¿¨
																									EM4305CopyEM4100(Buff);
																									ReadWriteEM4305(ReadCom,0,0,64);
																								break;
																			}
																			Carrier_Off();	
																			EM_Wait = 100;
																			while(EM_Wait > 0);
																			
																			Carrier_On();	 				//´ò¿ª´Å³¡	
																			EM_Wait = 1000;
																			Uart_Buff[1] = 0xF4;
																			Uart_Buff[2] = 0;
																			Uart_Len = 3;	
																			while(EM_Wait > 0)
																			{
																					if(ReadEM41xxCardNo())   //¶Á³öÀ´±È½Ï
																					{
																							for(i = 0; i < 5; i ++)
																							{
																									if(ID[i] != ListEM4100UID[i])  break;
																							}
																							if(i == 5)
																							{
																									Uart_Buff[2] = 1;
																									BP_BP(1,0);
																									for(i = 0; i < 5; i ++) Uart_Buff[Uart_Len ++] = ID[i];
																							}
																							break;
																					}
																			}
																			if(Uart_Buff[2] == 0)        //´íÎó
																			{
																					BP_BP(2,0);
																					for(i = 0; i < 5; i ++) Uart_Buff[Uart_Len ++] = 0;																			
																			}
																			CRC16_Base._uint = GetCRC16(Uart_Buff,Uart_Len); 	
																			Uart_Buff[Uart_Len ++] = CRC16_Base.B[0];
																			Uart_Buff[Uart_Len ++] = CRC16_Base.B[1];
																			Uart_Buff[Uart_Len ++] = 0xEE;
																			UartWrite(Uart_Buff,Uart_Len);	 //·µ»ØÊý¾Ý																					
																			break;
														case 0xF5:				//ÐÝÃß¿¨Æ¬
																			ReadWriteEM4305(DisableCom,0,BankData.uL,Read_PLL);
																			BP_BP(1,0);
																			break;
													  default:
																			break;
												}
										}	
										LED = 0;
								}
						}
				}
		}
}

void TM3_Isr() interrupt 19 using 1
{
		AUXINTIF &= ~0x02;              //Çå³ýÖÐ¶Ï±êÖ¾
	
	  Uart1RxMonitor(1);
	  if(LED_flag)	Time_Up ++;
	  if(Time_Up >= 500)
		{
				Time_Up = 0;
			  LED1 = ~LED1;					
		}
		if(EM_Wait > 0) EM_Wait --;
}
