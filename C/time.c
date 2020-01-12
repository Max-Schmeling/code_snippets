#include <time.h>

void delay(unsigned int milliseconds)
/*
 * Delay the program for a time of <milliseconds>
 * Requires: <time.h>
 */
{
    clock_t goal = clock() + milliseconds;
    while (goal > clock()) 
    	;
}
