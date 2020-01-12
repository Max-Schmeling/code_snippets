#include <stdlib.h>

int higher(int a, int b)
/*
 * Returns the higher number of a and b
 */
{
	return a > b ? a : b;
}


int lower(int a, int b)
/*
 * Returns the lower number of a and b
 */
{
	return a < b ? a : b;
}


int min(int *numbers, int len)
/*
 * Takes a number array as an argument
 * and returns the lowest value.
 */
{
	int lowest = numbers[0];
	for (int n = 1; n < len; n++) {
		if (numbers[n] < lowest) {
			lowest = numbers[n];
		}
	}
	return lowest;
}



int max(int *numbers, int len)
/*
 * Takes a number array as an argument
 * and returns the highest value.
 */
{
	int heighest = numbers[0];
	for (int n = 1; n < len; n++) {
		if (numbers[n] > heighest) {
			heighest = numbers[n];
		}
	}
	return heighest;
}



double mean(int *numbers, int len)
/*
 * Takes a number array as an argument
 * and returns the average of all values.
 */
{
	int sum = 0;
	for (int n = 0; n < len; n++) {
		sum += numbers[n];
	}
	return (double) sum / len;
}



double median(int *numbers, int len)
/*
 * Takes a number array as an argument
 * and returns the median. Ie. value in the
 * center. If <len> is even the median
 * is the average/mean of the center values.
 */
{
	// requires sorting beforehand
	//...

	if (len % 2 == 0) { // median is average of center values
		return (numbers[(int) (len / 2)] + numbers[(int) (len / 2 - 1)]) / 2.0;
	} else { // median is center value
		return numbers[(int) len / 2];
	}
}



long long power(int base, int exp)
/*
 * Calculates the power of base and (exp)onent.
 * Only allows whole-numbered exponents!
 */
{
	if (exp == 0) {
		return 1;
	}
	
	long long result = 1;
	for (int i = 1; i <= exp; i++)
	{
		result = result * base;
	}
	return result;
}


long faculty(int number)

{
	long result = 1;
	for (int i = number; i >= 1; i--) {
		result *= i;
	}
	return result;
}



void findCommonDivisor(long *numbers, int len)
/*
 * Returns commmon divisors of integers/longs
 * in array <numbers>. If there is only one
 * number (len = 1) print the factor pairs.
 */
{
	int count = 0; // counts factors found
	int flag = 1;	// 0 means that there is at least 
					// one number in <numbers> that
					// is not remainderlessly devisable.

	for (long i = 1; i <= min(numbers, len); i++) {
		for (int n = 0; n < len; n++) {
			if (numbers[n] % i != 0) {
				flag = 0;
				break;
			}
		}
		if (flag) {
			if (len == 1) {
				if (i <= numbers[0]/i) {
					printf("\n%d (x %d)", i, numbers[0]/i);
				} else {
					return;
				}
			} else {
				printf("\n%d", i);
			}
		}
		flag = 1;
	}
}


void findCommonDivisor2(int a, int b)
/*
 * Returns commmon divisors of a and b.
 * Can be used to get divisors of one
 * number by setting a = b.
 */
{
	int count = 0;

	for (int i = 1; i <= (a > b ? a : b); i++)
	{
		if ((a % i == 0) && (b % i == 0))
		{
			count++;
			printf("%d: %d\n", count, i);
		}
	}
}