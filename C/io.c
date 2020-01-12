int FileExist()
/* 
 * Create storage file if it does not exist.
 * Returns: 0 if file created/exists. 1 otherwise.
 */
{
	FILE *file;
	if (file = fopen("FILEPATH", "r")) {
		fclose(file);
		return 1;
	}
	return 0;
}