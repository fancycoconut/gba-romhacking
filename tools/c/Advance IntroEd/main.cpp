/*	--------------------------------------------
	Advance IntroEd
	By ZodiacDaGreat
	-------------------------------------------- */
#include <stdio.h>
#include <windows.h>

#include "nicestuff.h"
#include "resource.h"

char *FilePath;
char Header[4];

FILE *fp;

void LoadROM(HWND hWnd, char *File);
BOOL CALLBACK DialogProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nShowCmd)
{
	DialogBox(hInstance, MAKEINTRESOURCE(dlgMain), NULL, DialogProc);

	return 0;
}

BOOL CALLBACK DialogProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	HICON hIcon = NULL;

	switch(message)
	{
		case WM_INITDIALOG:
			hIcon = (HICON)LoadImage(GetModuleHandle(NULL), MAKEINTRESOURCE(appIcon), IMAGE_ICON, 16, 16, 0);
			SendMessage(hWnd, WM_SETICON, ICON_SMALL, (LPARAM)hIcon);

			break;

		case WM_COMMAND:
			switch(LOWORD(wParam))
			{
				case mnuExit:
					EndDialog(hWnd, 0);
					break;

				case mnuOpen:
					FilePath = ShowOpen(hWnd, "Open ROM...", "Gameboy Advance ROMs (*.gba;*.agb;*.bin)\0*.gba;*.agb;*.bin");
					if (FilePath != 0)
					{			
						LoadROM(hWnd, FilePath);
					}
					break;

				case IDCANCEL:
					EndDialog(hWnd, 0);
					break;
			}
			break;

		case WM_DESTROY:
			DestroyIcon(hIcon);
			break;
	}
	return 0;
}

void LoadROM(HWND hWnd, char *File)
{
	fp = fopen(File, "r+b");
	fseek(fp, 0xAC, SEEK_SET);
	fread(Header, 4, 1, fp);
	fclose(fp);

	switch()
	{
		case 0:
			MessageBox(hWnd, "1", "OK", MB_OK);
			break;

		case 2:
			MessageBox(hWnd, "2", "OK", MB_OK);
			break;

		default:
			MessageBox(hWnd, "Error - Unsupported ROM\nSupported versions are AXVE, BPRE, & BPEE :P", "Advance IntroEd", MB_ICONEXCLAMATION);
			break;
	}
	MessageBox(hWnd, Header, "OK", MB_OK);
}