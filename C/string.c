#include <stdio.h>
#include <string.h>
#include <ctype.h>


char *strrev(char *str)
/*
 * Returns reversed version of <str>
 */
{
    if (!str || ! *str)
        return str;

    int i = strlen(str) - 1, j = 0;

    char ch;
    while (i > j)
    {
        ch = str[i];
        str[i] = str[j];
        str[j] = ch;
        i--;
        j++;
    }
    return str;
}


int isSubstring(char *string, char *substring)
/* 
 * SYNOPSIS:
 *  Checks for substring in string, case insensitive.
 *  Does not allow german "umlaute" (ie. öäüß).
 *
 * RETURNS:
 *  1 if <substring> is substring of <string>, else 0.
 *
 * PROCEDURE:
 *  Iterates through every character of a string and
 *  compares the first one
 */
{
	int substrlen = strlen(substring); // For efficiency
	int i, flag = 0;
	long chars = 0;

	// Convert <substring> to lowercase before actual
	// comparison for case insensitivity and efficiency.
	for(i = 0; substring[i]; i++) {
  		substring[i] = tolower(substring[i]);
	}

	for (long chr = 0; chr < strlen(string); chr++) {
		chars++;
		if (tolower(string[chr]) == substring[flag]) {
			if (flag < substrlen) {
				if (flag+1 == substrlen) {
					return 1; // True
				}
				flag++;
			}
		} else {
			flag = 0;
		}
	}
	return 0; // False
}


int startswith(const char *str, const char *startstr)
/*
 * Description:
 *  Checks if <string> starts with <startstring>.
 *
 * Principle:
 *  Iterates through every character of <startstring>
 *  and checks if the current character is the same
 *  as the current one of <string>. If this is wrong
 *  for at least one character, return false. Ie.
 *  <string> does not start with <startstring>
 *
 * Returns: 
 *  1 when True; 0 when False or error
 */
{
	int c;
	
	if (strlen(str) < strlen(startstr)) {
		return 0; // False
	}
	
	for (c = 0; c < strlen(startstr); c++) {
		if (startstr[c] != str[c]) {
			return 0; // False
		}
	}
	return 1; // True
}


int endswith(char *string, char *endstring)
/*
 * Description:
 *  Checks if <string> ends with <endstring>.
 *
 * Principle:
 *  Iterates through every character of <endstring>
 *  in reverse order and checks if the current char
 *  is the same as the current one of <string>. If
 *  this is wrong for one character, the condition
 *  is false.
 *
 * Returns: 
 *  1 when True; 0 when False or error
 */
{
	int c, i = 0;
	int stringlen = strlen(string);
	int endstringlen = strlen(endstring);
	
	if (stringlen < endstringlen) {
		return 0; // Error
	}
	
	for (c = endstringlen-1; c >= 0; c--) {
		i++;
		if (string[stringlen-i] != endstring[c]) {
			return 0; // False
		}
	}
	return 1; // True
}


int isStrFloat(const char *str, int strlen)
/*
 * Returns 1 if a string is a floating point number.
 * Criteria: A string that must contain at least one
 * digit and one dot (ie. '.') at max. That is it does
 * not need to contain a dot
 */
{
	if (strlen == 0)
		return 1; // False: Empty string
	else if (strlen == 1 && str[0] == 46)
		return 1; // False: Only character is a dot
	
	int i, dotcount = 0;
	for (i = 0; i < strlen; i++) {
		if (str[i] == 46) {
			dotcount++;
			if (dotcount > 1)
				return 1; // False: More than one dot in string
		}
		else if (str[i] < 48 || str[i] > 57) {
			return 1; // False: Not a digit
		}
	}
	return 0; // True
}

int isStrInt(const char *str)
/*
 * Returns 1 if a string is a whole number.
 * Ie. It exclusively contains digits between
 * 0 and 9.
 */
{
	for (int d = 0; d != '\0'; d++) {
		if (str[d] < 48 || str[d] > 57) {
			return 0; // str is not a whole number
		}
	}
	return 1; // string is a whole number
}

int isStrBin(const char *str)
/*
 * Returns 1 if a string is binary.
 * Ie. it only contains 1s and 0s.
 */
{
	int b = 0;
	if (startswith(str, "0b")) {
		b = 2;
	} else if (startswith(str, "b")) {
		b = 1;
	}

	for (b; str[b] != '\0'; b++) {
		if (str[b] < '0' || str[b] > '1') {
			return 0; // False
		}
	}
	return 1; // True
}

int isStrHexnum(const char *str)
/* 
 * Returns 1 if a string is a
 * hexadecimal number
 */
{
	int c = 0;
	if (startswith(str, "0x")) {
		c = 2;
	}

	for (c; str[c] != '\0'; c++) {
		if (!((str[c] >= 'A' && str[c] <= 'F') || (str[c] >= '0' && str[c] <= '9'))) {
			return 0; // False
		}
	}
	return 1;
}

int isStrOctal(const char *str)
/*
 * Returns 1 if a string is an
 * octal number
 */
{
	for (int c = 0; str[c] != '\0'; c++) {
		if (str[c] < '0' || str[c] > '7') {
			return 0; // False
		}
	}
	return 1; // True
}




int wcmatch(const char *pattern, const char *string, int p, int c) {
	/*
	 * Checks if wildcard <pattern> matches <string> using recursion.
	 * <p> and <c> both need to be 0 on first call.
	 * Returns 1 if true and 0 if false.
	 */
	if (pattern[p] == '\0') {
		return string[c] == '\0';
	} else if (pattern[p] == '*') {
    	for (; string[c] != '\0'; c++) {
      		if (wcmatch(pattern, string, p+1, c))
        		return 1; // True
    	}
    	return wcmatch(pattern, string, p+1, c);
	} else if (pattern[p] != '?' && pattern[p] != string[c]) {
		return 0; // False
	} else {
		return wcmatch(pattern, string, p+1, c+1);
	}
}

