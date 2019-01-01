#include <windows.h>

int main()
{

    return SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, 0, SPIF_UPDATEINIFILE) - 1 ;
}


