double cels2fahr(double celsius) {
	return (celsius * (9.0/5.0)) + 32.0;
}

double cels2kelv(double celsius) {
	return celsius + 273.15;
}

double fahr2cels(double fahrenheit) {
	return ((fahrenheit- 32.0) * (5.0/9.0));
}

double fahr2kelv(double fahrenheit) {
	return (fahrenheit - 32) * (5.0/9.0) + 273.15;
}

double kelv2cels(double kelvin) {
	return kelvin - 273.15;
}

double kelv2fahr(double kelvin) {
	return kelvin * (9.0/5.0) - 459.67;
}