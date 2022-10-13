#include <stdio.h>
#include <windows.h>
#include <commctrl.h>

#include "resource.h"
#include "nicestuff.h"

BOOL CALLBACK DialogProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nShowCmd)
{
	DialogBox(hInstance, MAKEINTRESOURCE(frmMain), NULL, DialogProc);

	return 0;
}

BOOL CALLBACK DialogProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch(message)
	{
		case WM_INITDIALOG:
			break;

		case BN_CLICKED:
			break;

		case BN_DBLCLK:
			break;

		case WM_COMMAND:
			switch(LOWORD(wParam))
			{
				case mnuOpenTileset:
					BITMAP bm;
					HBITMAP Tileset;
					char *TilesetPath;

					TilesetPath = ShowOpen(hwnd, "Open Tileset...", "Bitmaps (*.bmp)\0*.bmp");
					if (TilesetPath != 0) // if lenb(TilesetPath) = 0 exit sub
					{
						Tileset = (HBITMAP)::LoadImage(NULL, TilesetPath, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE);
						GetObject(Tileset, sizeof(bm), & bm);

						HDC TilesetHDC = CreateCompatibleDC(TilesetHDC);
						SelectObject(TilesetHDC, Tileset);
						BitBlt(GetDC(GetDlgItem(hwnd, picTilebox)), 0, 0, bm.bmWidth, bm.bmHeight, TilesetHDC, 0, 0, SRCCOPY);
					}
					break;

				case mnuQuit:
					EndDialog(hwnd, 0);
					break;

				case IDCANCEL:
					EndDialog(hwnd, 0);
					break;
			}
			break;

		case WM_LBUTTONDBLCLK:
			break;

		case WM_DESTROY:
			break;
	}

	return 0;
}