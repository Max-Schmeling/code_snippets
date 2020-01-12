void dec2bin(long number, char *binary)
/*
 * Converts a number of base 10 (<decimal>)
 * to a binary string (<binary>).
 */
{
	int i = 0, j;

	for (number; number > 0; number/=2) {
		if (number % 2 == 0) {
			binary[i++] = '0';
		} else if (number % 2 == 1) {
			binary[i++] = '1';
		}
	}
	binary[i] = '\0';

	// Reverse binary string
	strrev(binary);
	/*
	for (j = 0; j < (i/2); j++) {
		printf("\n%d : %c -> %c", j, binary[j], binary[i - j - 1]);
		binary[j] = binary[i - 1 - j];
	}
	*/
}


long dec2bin2(long decimal)
/*
 * Converts a number of base 10
 * to binary.
 *
 * CAUTION:
 * Supports maximum decimal of 1023
 * because <b> reaches LONG_MAX.
 */
{
	if (decimal > 1023) { return 0; }

	long binary = 0;
	long b = 1;
	while (decimal != 0) {
		binary += decimal % 2 * b;
		decimal /= 2;
		b *= 10;
	}
	return binary;
}


void dec2hex(long decimal, char *hexstring)
{
	sprintf(hexstring, "%X", decimal);
}


void dec2oct(long decimal, char *octstring)
{
	sprintf(octstring, "%o", decimal);
}


long bin2dec(char bin[])
/*
 * Converts a given binary string to an integer
 */
{
	int b;
	long decimal = 0;
	for (b = 0; bin[b] != '\0'; b++) {
		if (bin[b] == '1') {
			decimal = decimal * 2 + 1;
		} else if (bin[b] == '0') {
			decimal *= 2;
		}
	}
	return decimal;
}


unsigned long hex2dec(char hex[])
/*
 * Returns the integer representation of a hex value
 */
{
	return strtoul(hex, NULL, 16);
}


unsigned long oct2dec(char octal[])
/*
 * Converts an octal string to an integer
 */
{
	return strtoul(octal, NULL, 8);
}