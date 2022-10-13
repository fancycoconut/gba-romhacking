#include <windows.h>

OPENFILENAME cdlg;

char *ShowOpen(HWND hWnd, const char *Title, const char *FilterList)
{
	char *File = "";

	ZeroMemory(&cdlg, sizeof(cdlg));

	cdlg.lStructSize = sizeof(cdlg);
	cdlg.hwndOwner = hWnd;
	cdlg.lpstrTitle = Title;
	cdlg.lpstrFilter = FilterList;
	cdlg.nMaxFile = MAX_PATH;
	cdlg.lpstrFile = File;
	cdlg.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;

	if (GetOpenFileName(&cdlg))
	{
		ZeroMemory(&cdlg, sizeof(cdlg));
		return File;
	}
}

char *ShowSave(HWND hWnd, const char *Title, const char *FilterList)
{
	char *File = "";

	ZeroMemory(&cdlg, sizeof(cdlg));

	cdlg.lStructSize = sizeof(cdlg);
	cdlg.hwndOwner = hWnd;
	cdlg.lpstrTitle = Title;
	cdlg.lpstrFilter = FilterList;
	cdlg.nMaxFile = MAX_PATH;
	cdlg.lpstrFile = File;
	cdlg.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;

	if (GetSaveFileName(&cdlg))
	{
		ZeroMemory(&cdlg, sizeof(cdlg));
		return File;
	}
}