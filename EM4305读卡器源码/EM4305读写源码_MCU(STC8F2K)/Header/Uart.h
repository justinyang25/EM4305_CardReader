#ifndef __Uart_H__
#define __Uart_H__

extern bit cmdArrived;
extern unsigned char idata bufRxd[50];  //Ω” ’ª∫≥Â
extern void UartInit(void);
extern void Uart1RxMonitor(unsigned char ms);
extern unsigned char UartRead(unsigned char *buf, unsigned char len); 
extern void UartWrite(unsigned char *buf, unsigned char len);
//extern void UartSend(unsigned char dat);	

#endif