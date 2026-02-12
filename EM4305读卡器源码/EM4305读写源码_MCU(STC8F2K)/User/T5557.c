#include"T5557.h"
#include"EM41xx.h"

#define    Start_Gap  200   	   //起始   160 200 	   40*8=320
#define    GapS       100        //间隔   60  150	     15*8=120
#define    Bit_One    220    	   //数据1  220 350	     55*8=440 
#define    Bit_Zero   100    	   //数据   100 100      25*8=200


void DelayNus(unsigned int us)	 //@24MHz
{
		unsigned char i;
	
		while(us--)
		{
				i = 6;
	  		while (--i);
		}
}

void Contrl_RF(bit state,unsigned int dtime)	//控制载波时长 
{
		if(state) 
		{
				Carrier_On();			   //开启载波
		}
		else
		{
				Carrier_Off();			   //关闭载波
		}
		DelayNus(dtime);	
}

void Standard_Write(unsigned char Block,unsigned long W_data)
{
		unsigned char i;
		unsigned char Tbit;
	
		Carrier_On();			  	  //开启载波
		DelayNus(10000);		    //POR 启动延时6MS 
		Contrl_RF(0,Start_Gap);

		Contrl_RF(1,Bit_One); 	//10 对0页  操作码
		Contrl_RF(0,GapS);
		Contrl_RF(1,Bit_Zero);  //0
		Contrl_RF(0,GapS);

		Contrl_RF(1,Bit_Zero);  //块锁信号 0：不锁
		Contrl_RF(0,GapS);

		for(i=0;i<32;i++)
		{
				W_data = W_data<<1;
				if(CY)					       				//W_data&Tbit
				{
						Contrl_RF(1,Bit_One);    	//1
				}
				else
				{
						Contrl_RF(1,Bit_Zero);   	//0
				}
				Contrl_RF(0,GapS);   					//数据间隔			
		}									
		Tbit = 0x04;											//块地址 0000 0100
		for(i=0;i<3;i++)
		{		
				if(Block & Tbit)
				{
						Contrl_RF(1,Bit_One);    	//1	
				}
				else
				{
						Contrl_RF(1,Bit_Zero);    //0
				}
				Contrl_RF(0,GapS);   					//数据间隔
				Tbit >>= 1;										//0000 0010
		}
		Carrier_On();										  //给卡一个编程的时间
		DelayNus(10000);
		Carrier_Off();
}