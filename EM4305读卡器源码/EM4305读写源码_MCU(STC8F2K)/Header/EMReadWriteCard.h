#ifndef __EMReadWriteCard_H_
#define __EMReadWriteCard_H_
#include"STC8F.h"
#include"IntPut_IO.h"
#include"EM41xx.h"

#define LoginCom    0x03         //µÇÂ¼ÃüÁî
#define WriteCom    0x05         //Ğ´ÃüÁî
#define ReadCom     0x09         //¶ÁÃüÁî
#define ProtectCom  0x0C         //±£»¤ÃüÁî
#define DisableCom  0x0A         //ĞİÃßÃüÁî


extern unsigned char EM_Buff[10];
extern void Delay_8nus(unsigned int n);
unsigned char ReadWriteEM4305(unsigned char _cardCom,unsigned char _block,unsigned long _blockData,unsigned char _nRF);

#endif